VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{FBAFE9A8-8B26-4559-9D12-D70E36A97BE3}#2.1#0"; "zlRichEditor.ocx"
Begin VB.Form frmSentenceEdit 
   Caption         =   "词句示范编辑"
   ClientHeight    =   5535
   ClientLeft      =   2775
   ClientTop       =   3765
   ClientWidth     =   8100
   Icon            =   "frmSentenceEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5535
   ScaleWidth      =   8100
   StartUpPosition =   1  '所有者中心
   Begin MSComctlLib.TreeView tvw分类 
      Height          =   2430
      Left            =   1335
      TabIndex        =   17
      Tag             =   "1000"
      Top             =   1500
      Visible         =   0   'False
      Width           =   4950
      _ExtentX        =   8731
      _ExtentY        =   4286
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "imgList"
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.TextBox txt分类 
      Height          =   300
      Left            =   1320
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   12
      Text            =   "(无)"
      Top             =   1200
      Width           =   4980
   End
   Begin VB.CommandButton cmd分类 
      Caption         =   "归类到(&L)"
      Height          =   350
      Left            =   105
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1170
      Width           =   1215
   End
   Begin VB.ComboBox cbo科室 
      Height          =   300
      Left            =   855
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   795
      Width           =   3660
   End
   Begin VB.PictureBox pic内容 
      Height          =   3660
      Left            =   30
      ScaleHeight     =   3600
      ScaleWidth      =   7950
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1605
      Width           =   8010
      Begin zlRichEditor.Editor edt内容 
         Height          =   3060
         Left            =   75
         TabIndex        =   14
         Top             =   420
         Width           =   7755
         _ExtentX        =   13679
         _ExtentY        =   5398
         PaperHeight     =   11907
         PaperWidth      =   16840
         WithViewButtonas=   0   'False
         PaperKind       =   4
         ShowRuler       =   0   'False
         AuditMode       =   -1  'True
      End
      Begin XtremeCommandBars.CommandBars cbsThis 
         Left            =   0
         Top             =   0
         _Version        =   589884
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
         VisualTheme     =   2
      End
   End
   Begin VB.OptionButton opt范围 
      Caption         =   "个人使用(&3)"
      Height          =   180
      Index           =   2
      Left            =   4140
      TabIndex        =   7
      Top             =   525
      Value           =   -1  'True
      Width           =   1305
   End
   Begin VB.OptionButton opt范围 
      Caption         =   "科内通用(&2)"
      Height          =   180
      Index           =   1
      Left            =   2497
      TabIndex        =   6
      Top             =   525
      Width           =   1305
   End
   Begin VB.OptionButton opt范围 
      Caption         =   "全院通用(&1)"
      Height          =   180
      Index           =   0
      Left            =   855
      TabIndex        =   5
      Top             =   525
      Width           =   1305
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6705
      TabIndex        =   15
      Top             =   105
      Width           =   1215
   End
   Begin VB.TextBox txt名称 
      Height          =   300
      Left            =   3075
      TabIndex        =   3
      Top             =   105
      Width           =   3225
   End
   Begin VB.TextBox txt编号 
      Height          =   300
      Left            =   855
      TabIndex        =   1
      Top             =   105
      Width           =   1260
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6705
      TabIndex        =   16
      Top             =   525
      Width           =   1215
   End
   Begin MSComctlLib.ImageList imgList 
      Bindings        =   "frmSentenceEdit.frx":058A
      Left            =   105
      Top             =   1125
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
            Picture         =   "frmSentenceEdit.frx":059E
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSentenceEdit.frx":0B38
            Key             =   "expend"
         EndProperty
      EndProperty
   End
   Begin VB.Label lbl科室 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "制作(&R)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   135
      TabIndex        =   8
      Top             =   855
      Width           =   630
   End
   Begin VB.Label lbl人员 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   4560
      TabIndex        =   10
      Top             =   795
      Width           =   1740
   End
   Begin VB.Label lbl范围 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "使用(&U)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   135
      TabIndex        =   4
      Top             =   525
      Width           =   630
   End
   Begin VB.Label lbl名称 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "名称(&N)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   2400
      TabIndex        =   2
      Top             =   165
      Width           =   630
   End
   Begin VB.Label lbl编号 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "编号(&D)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   135
      TabIndex        =   0
      Top             =   165
      Width           =   630
   End
End
Attribute VB_Name = "frmSentenceEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------
'说明：
'   1、上级程序通过本窗体ShowMe函数，将父窗体、编辑词句ID,编辑状态等信息传递进入本程序
'   2、编辑状态：由Me.tag存放，分别为"新增"、"修改"，由上级程序通过ShowMe传入
'---------------------------------------------------
Private mlngClassId As Long                         '分类ID
Private mlngWordId As Long                          '词句ID
Private mblnOK As Boolean                           '是否完成编辑退出

Private Elements As cEPRElements                    '局部诊治要素集合
Private WithEvents mfrmInsElement As frmInsElement  '插入诊治要素窗体
Attribute mfrmInsElement.VB_VarHelpID = -1
Private mlngHP As Long, blnSpaceEvent As Boolean    '记录自动增加空格的位置！

Private blnActive As Boolean

'临时变量
Dim lngCount As Long

'-----------------------------------------------------
'以下为外部公共程序
'-----------------------------------------------------
Public Function ShowMe(ByVal frmParent As Form, _
    ByVal blnAdd As Boolean, ByVal bytPower As Byte, ByVal lngClassId As Long, _
    Optional ByVal lngWordId As Long, Optional ByVal blnSaveAs As Boolean) As Long
    '---------------------------------------------------
    '功能：上级程序调用本窗体的，传递参数，并显示窗体
    '参数： bytPower-管理权限（0-全院；1-科室；2-个人）
    '       lngClassId-分类id
    '       lngWordId-记录ID，修改时必须
    '       blnSaveAs-是否病历编辑过程中的“另存词句示范”调用
    '返回：确定返回新增或修改的ID；取消返回0
    '---------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    blnActive = False
    mlngClassId = lngClassId: mlngWordId = lngWordId
    If blnAdd Then
        Me.Tag = "新增": mlngWordId = 0
    Else
        Me.Tag = "修改"
    End If
    
    '---------------------------------------------------
    '基本数据信息
    Dim objNode As MSComctlLib.Node
    Err = 0: On Error GoTo ErrHand
    
    gstrSQL = "Select ID, 上级id, 编码, 名称, 说明" & vbNewLine & _
            "From 病历词句分类" & vbNewLine & _
            "Start With 上级id Is Null" & vbNewLine & _
            "Connect By Prior ID = 上级id" & vbNewLine & _
            "Order By Level, 编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    With rsTemp
        Me.tvw分类.Nodes.Clear
        Do While Not .EOF
            If IsNull(!上级ID) Then
                Set objNode = Me.tvw分类.Nodes.Add(, , "_" & !ID, !编码 & "-" & !名称, "close")
            Else
                Set objNode = Me.tvw分类.Nodes.Add("_" & !上级ID, tvwChild, "_" & !ID, !编码 & "-" & !名称, "close")
            End If
            objNode.Sorted = True
            objNode.ExpandedImage = "expend"
            If !ID = lngClassId Then
                objNode.Selected = True
                Me.txt分类.Tag = !ID: Me.txt分类.Text = objNode.Text
            End If
            .MoveNext
        Loop
    End With
    
    gstrSQL = "Select Distinct d.Id, d.编码, d.名称, r.缺省, r.人员id, p.姓名 " _
            & "From 部门表 d, 部门人员 r, 人员表 p, 上机人员表 u, 部门性质说明 c " _
            & "Where d.Id = r.部门id And r.人员id = p.Id And p.Id = u.人员id And u.用户名 = User And d.Id = c.部门id And " _
            & "      c.工作性质 In ('临床', '检查', '检验', '手术', '治疗', '护理', '营养', '体检') And (p.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or p.撤档时间 Is Null) " _
            & "Order By r.缺省 Desc,d.编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    With rsTemp
        Me.cbo科室.Clear
        Do While Not .EOF
            Me.cbo科室.AddItem !编码 & "-" & !名称
            Me.cbo科室.ItemData(Me.cbo科室.NewIndex) = !ID
            If !缺省 = 1 Then Me.cbo科室.ListIndex = Me.cbo科室.NewIndex
            Me.lbl人员.Tag = !人员ID: Me.lbl人员.Caption = !姓名
            .MoveNext
        Loop
        If Me.cbo科室.ListCount = 0 Then
            MsgBox "你目前不属于任何临床/检查/检验/手术/治疗/护理/营养/体检部门，不能管理词句示范！", vbExclamation, gstrSysName
            ShowMe = 0: Unload Me: Exit Function
        ElseIf Me.cbo科室.ListIndex = -1 Then
            Me.cbo科室.ListIndex = 0
        End If
    End With
    
    '---------------------------------------------------
    '内容数据提取
    gstrSQL = "Select l.分类id, l.编号, l.名称, l.通用级, l.科室id, d.编码, d.名称 As 部门, l.人员id, p.姓名 As 人员 " _
            & "From 病历词句示范 l, 部门表 d, 人员表 p " _
            & "Where l.科室id = d.Id And l.人员id = p.Id And l.id =[1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngWordId)
    With rsTemp
        If .RecordCount > 0 Then
            Me.txt编号.Text = !编号
            Me.txt名称.Text = !名称
            Me.tvw分类.Nodes("_" & !分类id).Selected = True
            Me.txt分类.Tag = !分类id: Me.txt分类.Text = Me.tvw分类.SelectedItem.Text
            Me.opt范围(IIf(IsNull(!通用级), 0, !通用级)).Value = True
            If !人员ID <> Me.lbl人员.Tag Then
                Me.lbl人员.Tag = !人员ID: Me.lbl人员.Caption = !人员
                Me.cbo科室.Clear
                Me.cbo科室.AddItem !编码 & "-" & !部门
                Me.cbo科室.ItemData(Me.cbo科室.NewIndex) = !科室ID
                Me.cbo科室.ListIndex = Me.cbo科室.NewIndex
                Me.cbo科室.Enabled = False
            Else
                For lngCount = 0 To Me.cbo科室.ListCount - 1
                    If Me.cbo科室.ItemData(lngCount) = IIf(IsNull(!科室ID), 0, !科室ID) Then
                        Me.cbo科室.ListIndex = lngCount: Exit For
                    End If
                Next
            End If
        End If
        Me.txt编号.MaxLength = .Fields("编号").DefinedSize
        Me.txt名称.MaxLength = .Fields("名称").DefinedSize
    End With
    
    If InStr(1, gstrPrivsEpr, "全院病历词句") <> 0 Then
        
    ElseIf InStr(1, gstrPrivsEpr, "科室病历词句") <> 0 Then
        Me.opt范围(0).Enabled = False
    ElseIf InStr(1, gstrPrivsEpr, "个人病历词句") <> 0 Then
        Me.opt范围(0).Enabled = False: Me.opt范围(1).Enabled = False
    End If
    If Me.Tag = "新增" Then Call zlDefaultCode
    
    '---------------------------------------------------
    '词句内容恢复
    If blnAdd = False Then
        Call InsertPhrase(mlngWordId)
    ElseIf blnSaveAs Then
        Call InsertSelText(frmParent)
    End If
    
    '---------------------------------------------------
    Call InitMenu
    '---------------------------------------------------
    '显示窗体
    Me.edt内容.AuditMode = False
    If frmParent Is Nothing Then
        Me.Show vbModal
    Else
        Me.Show vbModal, frmParent
    End If
    If mblnOK Then
        ShowMe = mlngWordId
    Else
        ShowMe = 0
    End If
    Unload Me
    Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    ShowMe = 0
End Function

'-----------------------------------------------------
'以下为内部公共程序
'-----------------------------------------------------
Private Sub zlDefaultCode()
    '功能：设置默认的编号
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "Select LPad(Nvl(To_Number(Max(编号)), 0) + 1, Nvl(Max(Length(编号)), 5), '0') As 编码" & vbNewLine & _
            "From 病历词句示范" & vbNewLine & _
            "Where 分类id = [1]"
    Err = 0: On Error Resume Next
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(Me.txt分类.Tag))
    Me.txt编号.Text = rsTemp.Fields(0).Value
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Me.txt编号.Text = ""
End Sub

Private Sub InitMenu()
    '功能： 编辑器工具栏设置
    Dim rsTemp As New ADODB.Recordset
    Dim cbrControl As CommandBarControl, cbrCombox As CommandBarComboBox
    '---------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    Set Me.cbsThis.Icons = zlCommFun.GetPubIcons
    With Me.cbsThis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = False
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    
    '-----------------------------------------------------
    '菜单定义
    Me.cbsThis.ActiveMenuBar.Title = "菜单"
    Me.cbsThis.ActiveMenuBar.EnableDocking xtpFlagAlignTop Or xtpFlagStretched Or xtpFlagHideWrap
    With cbsThis.ActiveMenuBar.Controls
        Set cbrControl = .Add(xtpControlLabel, 0, "可导入词句示范")
        Set cbrCombox = .Add(xtpControlComboBox, conMenu_Edit_Import, "示范列表")
        cbrCombox.Width = 160: cbrCombox.DropDownWidth = 180: cbrCombox.DropDownListStyle = True
        gstrSQL = "Select Id, 编号, 名称 From 病历词句示范 Where Id <> [1] And 分类id = [2] Order By 编号"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngWordId, mlngClassId)
        Do While Not rsTemp.EOF
            cbrCombox.AddItem rsTemp!编号 & "-" & rsTemp!名称
            cbrCombox.ItemData(rsTemp.AbsolutePosition) = rsTemp!ID
            rsTemp.MoveNext
        Loop
        cbrCombox.ListIndex = 1
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "插入要素(Ctrl+I)"): cbrControl.flags = xtpFlagRightAlign: cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改要素(Ctrl+M)"): cbrControl.flags = xtpFlagRightAlign
    End With
    For Each cbrControl In cbsThis.ActiveMenuBar.Controls
        If cbrControl.Type = xtpControlButton Then cbrControl.Style = xtpButtonIconAndCaption
    Next
    
    With Me.cbsThis.KeyBindings
        .Add FCONTROL, Asc("I"), conMenu_Edit_NewItem
        .Add FCONTROL, Asc("M"), conMenu_Edit_Modify
        .Add FCONTROL, Asc("J"), ID_INSERT_AUTORECOGNISE
        .Add 0, VK_DELETE, conMenu_Edit_Delete
        
        .Add FCONTROL, Asc("X"), ID_EDIT_CUT
        .Add FCONTROL, Asc("C"), ID_EDIT_COPY
        .Add FCONTROL, Asc("V"), ID_EDIT_PASTE
        .Add 0, VK_DELETE, conMenu_Edit_Delete
        .Add 0, VK_F12, ID_INSERT_AUTORECOGNISE              '智能识别
    End With
End Sub

Private Sub ValidteRTF()
    '清除RTF中的特殊文本和关键字
    Dim sType As String, lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bInKeys As Boolean, bNeeded As Boolean
    Dim i As Long, bFinded As Boolean
    i = 1
    
    Me.edt内容.ForceEdit = True
    Do
        bFinded = FindNextAnyKey(edt内容, i + 1, sType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
        If bFinded Then
            If sType <> "E" Then
                Me.edt内容.Range(lKSS, lKEE).Font.Protected = False
                Me.edt内容.Range(lKSS, lKEE).Font.ForeColor = tomAutoColor
                Me.edt内容.Range(lKSS, lKEE).Font.BackColor = tomAutoColor
                Me.edt内容.Range(lKSS, lKEE).Font.Hidden = False
                Me.edt内容.Range(lKES, lKEE) = ""
                Me.edt内容.Range(lKSS, lKSE) = ""
            Else
                i = lKEE
            End If
        Else
            i = i + 1
        End If
    Loop Until bFinded = False
    Me.edt内容.ForceEdit = False
End Sub

Private Sub AppendElement(ByRef Ele As cEPRElement)
    '添加要素
    Dim lngKey As Long, lngLen As Long
    lngLen = Len(Me.edt内容.Text)
    Me.edt内容.Range(lngLen, lngLen).Selected
    lngKey = Elements.AddExistNode(Ele, True)
End Sub

Private Sub AppendText(ByVal strText As String)
    '添加文本
    Dim lngKey As Long, lngLen As Long
    lngLen = Len(Me.edt内容.Text)
    Me.edt内容.ForceEdit = True
    Me.edt内容.Range(lngLen, lngLen) = strText
    Me.edt内容.ForceEdit = False
End Sub

'################################################################################################################
'## 功能：  显示自动识别诊治要素或者字典项目的选择器
'##
'## 参数：  strAuto     :IN     传入查询关键字
'################################################################################################################
Private Sub ShowAutoRecSelector(ByVal strF As String)
    Dim sType As String, lSS As Long, lSE As Long, lES As Long, lEE As Long, lKey As Long, bInKeys As Boolean, bNeeded As Boolean
    bInKeys = IsBetweenAnyKeys(edt内容, Me.edt内容.SelStart + 1, sType, lSS, lSE, lES, lEE, lKey, bNeeded)
    If bInKeys Then Exit Sub     '保证不能插入关键字内部
    If Me.edt内容.Selection.Font.Protected Then Exit Sub

    Dim rs As New ADODB.Recordset
    Dim lLeft As Long, lTOp As Long
    
    '如果中文名或者英文名建了索引会更快一些！
    gstrSQL = "select  ID,编码,中文名 As 名称,单位,decode(替换域,2,'字典项目',1,'替换项目','外部输入项') As 类型 " & _
        "From 诊治所见项目 " & _
        "Where 中文名 Like '%" & strF & "%' Or 英文名 Like '%" & UCase(strF) & "%' " & _
        "Order By 类型"
        
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If rs.EOF Then Exit Sub
    Dim pt As POINTAPI, arrPara As String, T As Variant, lngID As Long
    Dim f As New frmSelectChild
    
    pt.x = 0
    pt.y = 0
    ClientToScreen Me.edt内容.OriginRTB.hwnd, pt
    '获取起始位置坐标
    Me.edt内容.Range(edt内容.SelStart, Me.edt内容.SelStart + 1).GetPoint cprGPStart + cprGPLeft + cprGPBottom, lLeft, lTOp

    arrPara = "0;830;2500;700;1000"
    strF = f.ShowSelectChild(Me, pt.x * Screen.TwipsPerPixelX + lLeft, pt.y * Screen.TwipsPerPixelY + lTOp, _
        5550, 3000, rs, arrPara)
    If strF = "" Then
        Exit Sub
    Else
        T = Split(strF, ";")
        lngID = T(0)
        rs.Close
        gstrSQL = "Select ID, 中文名, 类型, 长度, 小数, 单位, 表示法, 替换域, 初始值, 数值域 From 诊治所见项目 Where ID =[1]"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngID)
        If Not rs.EOF Then
            '插入元素
            Dim Ele As New cEPRElement, aryTemp() As String, lngKey As Long, lngCount As Long
            With Ele
                .要素名称 = NVL(rs("中文名"))
                .诊治要素ID = NVL(rs("ID"), 0)
                .要素类型 = NVL(rs("类型"), 1)
                .要素长度 = NVL(rs("长度"), 0)
                .要素小数 = NVL(rs("小数"), 0)
                .要素单位 = NVL(rs("单位"))
                .要素表示 = IIf(NVL(rs("表示法"), 0) = 4, 2, NVL(rs("表示法"), 0))
                .替换域 = NVL(rs("替换域"), 0)      '0-外部输入项目；1-替换项目；2-字典项目
                .内容文本 = Trim(NVL(rs("初始值")))
                If .要素类型 = 0 Then
                    Select Case .要素表示
                    Case 0, 1
                        If Trim(NVL(rs("数值域"))) = "" Then
                            .要素值域 = ""
                        Else
                            aryTemp = Split(NVL(rs("数值域")), ";")
                            .要素值域 = Val(aryTemp(0)) & ";" & Val(aryTemp(1))
                        End If
                    Case 2
                        aryTemp = Split(NVL(rs("数值域")), ";")
                        For lngCount = 0 To UBound(aryTemp)
                            aryTemp(lngCount) = Val(aryTemp(lngCount))
                        Next
                        .要素值域 = Join(aryTemp(0), ";")
                    Case Else
                        .要素值域 = ""
                    End Select
                Else
                    Select Case .要素表示
                    Case 2, 3
                        .要素值域 = NVL(rs("数值域"))
                    Case Else
                        .要素值域 = ""
                    End Select
                End If
                .输入形态 = IIf(.要素表示 = 2 Or .要素表示 = 3, 1, 0) '0-文本 1-上下 2-单选 3-复选   如果为单选、复选，则这里默认值为展开项目   0-弹出;1-展开
            End With
            lngKey = Elements.AddExistNode(Ele)
            
            '插入诊治要素到编辑器中
            Dim blnForce As Boolean
            blnForce = Me.edt内容.ForceEdit
            Me.edt内容.ForceEdit = True
            Me.edt内容.SelText = ""
            Elements("K" & lngKey).InsertIntoEditor Me.edt内容, , True
            Me.edt内容.ForceEdit = blnForce
        End If
    End If
End Sub

'################################################################################################################
'## 功能：将指定词句示范插入编辑器，用于词句的修改内容恢复和其他词句插入
'################################################################################################################

Private Sub InsertPhrase(ByVal lngImpId As Long)
    Dim rsTemp As New ADODB.Recordset
    
    Dim sType As String, lSS As Long, lSE As Long, lES As Long, lEE As Long, lKey As Long, bInKeys As Boolean, bNeeded As Boolean
    Dim lngKey As Long, lngStart As Long, lngLen As Long, strTmp As String
    
    bInKeys = IsBetweenAnyKeys(edt内容, Me.edt内容.SelStart + 1, sType, lSS, lSE, lES, lEE, lKey, bNeeded)
    If bInKeys Then Exit Sub
    
    gstrSQL = "Select 词句id, 排列次序, 内容性质, 内容文本, 诊治要素id, 替换域, 要素名称, 要素类型, 要素长度, 要素小数, 要素单位, 要素表示, 要素值域, 输入形态, 对象属性" & vbNewLine & _
                "From 病历词句组成" & vbNewLine & _
                "Where 词句id = [1]" & vbNewLine & _
                "Order By 排列次序"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngImpId)
    With Me.edt内容
        .Freeze
        .ForceEdit = True
        Do While Not rsTemp.EOF
            Select Case rsTemp("内容性质")
            Case 0 '自由文字
                '恢复RTF内容
                lngStart = .SelStart
                strTmp = NVL(rsTemp("内容文本"))
                lngLen = Len(strTmp)

                .Range(lngStart, lngStart) = strTmp
                .Range(lngStart, lngStart + lngLen).Font.Protected = False
                .Range(lngStart, lngStart + lngLen).Font.Hidden = False
                .Range(lngStart + lngLen, lngStart + lngLen).Selected
            Case 1, 2 '1-临时诊治要素,2-固定诊治要素
                lngStart = .SelStart
                
                lngKey = Elements.Add
                Elements("K" & lngKey).ID = mlngWordId
                Elements("K" & lngKey).内容文本 = NVL(rsTemp("内容文本"))
                Elements("K" & lngKey).要素名称 = NVL(rsTemp("要素名称"))
                Elements("K" & lngKey).诊治要素ID = NVL(rsTemp("诊治要素ID"), 0)
                Elements("K" & lngKey).替换域 = NVL(rsTemp("替换域"), 0)
                Elements("K" & lngKey).要素类型 = NVL(rsTemp("要素类型"), 0)
                Elements("K" & lngKey).要素长度 = NVL(rsTemp("要素长度"), 0)
                Elements("K" & lngKey).要素小数 = NVL(rsTemp("要素小数"), 0)
                Elements("K" & lngKey).要素单位 = NVL(rsTemp("要素单位"))
                Elements("K" & lngKey).要素表示 = NVL(rsTemp("要素表示"), 0)
                Elements("K" & lngKey).要素值域 = NVL(rsTemp("要素值域"))
                Elements("K" & lngKey).输入形态 = NVL(rsTemp("输入形态"), 0)
                Elements("K" & lngKey).是否换行 = False
                Elements("K" & lngKey).对象属性 = NVL(rsTemp!对象属性)
                Elements("K" & lngKey).InsertIntoEditor Me.edt内容, lngStart, , True
            
            End Select
            rsTemp.MoveNext
        Loop
        lngStart = .SelStart
        .ForceEdit = False
        
        '去掉末尾的回车换行符
        lngLen = Len(Me.edt内容.Text)
        If lngLen > 0 Then
            If (Me.edt内容.Range(lngLen - 2, lngLen) = vbCrLf Or (Asc(Me.edt内容.Range(lngLen - 1, lngLen)) = 13 And Asc(Me.edt内容.Range(lngLen - 2, lngLen - 1)) = 10)) And Me.edt内容.Range(lngLen - 2, lngLen).Font.Protected = False Then
                Me.edt内容.Range(lngLen - 2, lngLen) = ""
            End If
        End If
        
        .Range(lngStart, lngStart).Selected
        .Modified = False
        .UnFreeze
    End With
End Sub

'################################################################################################################
'## 功能：将上级窗体的选中内容指定格式文本插入到编辑器中
'################################################################################################################
Private Sub InsertSelText(ByVal frmEdit As Form)
    Dim sType As String, lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bInKeys As Boolean, bNeeded As Boolean
    Dim lS As Long, lE As Long, i As Long, bFinded As Boolean
    Dim lngKey As Long, bBeteenKeys As Boolean
    
    '扩展保证完整的要素选择
    lS = frmEdit.Editor1.Selection.StartPos
    lE = frmEdit.Editor1.Selection.EndPos
    bBeteenKeys = IsBetweenAnyKeys(frmEdit.Editor1, lS + 1, sType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
    If bBeteenKeys Then lS = lKSS
    bBeteenKeys = IsBetweenAnyKeys(frmEdit.Editor1, lE + 1, sType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
    If bBeteenKeys Then lE = lKEE
    
    '先赋值RTF
    Me.edt内容.NewDoc
    Me.edt内容.ForceEdit = True
    Me.edt内容.TOM.TextDocument.Selection.FormattedText = frmEdit.Editor1.TOM.TextDocument.Range(lS, lE).FormattedText
    SetCommonStyle Me.edt内容, "正文", 0, Len(Me.edt内容.Text), True
    
    '处理要素
    Set Elements = New cEPRElements
    For i = lS To lE
        bFinded = FindNextAnyKey(frmEdit.Editor1, i + 1, sType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
        If bFinded = False Then Exit For    '不存在任何元素，那么退出循环
        If Not (lKSS < lE) Then Exit For    '超出范围，退出循环
        '范围内存在关键字
        If sType = "E" Then
            '如果是要素，那么拷贝到缓冲区
            Call AppendElement(frmEdit.Document.Elements("K" & lKey).Clone(True))
        End If
        i = lKEE - 1
    Next
    Call ValidteRTF
End Sub

'################################################################################################################
'## 功能：  获取纯文本到数据库的SQL语句
'##
'## 参数：
'##         ArraySQL()      :IN/OUT，   SQL数组
'##         strIn           :IN，       需要保存的字符串
'##         lng序号         :IN，       序号
'##         bln是否换行     :IN，       是否换行
'##
'## 说明：  长度大于4000的字符串，分行存储，序号递增之！
'################################################################################################################
Private Function GetPlainTextSaveSQL(ByRef ArraySQL() As String, _
    ByVal strIn As String, ByRef lng序号 As Long) As Boolean
    
    Dim lngLen As Long, strSub As String, i As Long, lngID As Long
    Dim lngCount As Long, lID As Long
    strIn = Replace(strIn, "'", "' || chr(39) || '")
    strIn = Replace(strIn, vbCrLf, "' || chr(13) || chr(10) || '")  '本来strIn是不允许有vbCrlf的。
    strIn = Replace(strIn, "", " ") '中文全角空格
    lngLen = Len(strIn)
    
    '按照4000为界分段存储。
    i = 0
    Do While (i * 2000 + 1 <= lngLen)
        lngCount = UBound(ArraySQL) + 1
        ReDim Preserve ArraySQL(1 To lngCount) As String

        strSub = Mid(strIn, i * 2000 + 1, 2000)

        gstrSQL = "Zl_病历词句组成_Insert(" & mlngWordId & "," & lng序号 & ",0,'" & strSub & "',NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL,NULL)"
        
        ArraySQL(lngCount) = gstrSQL
       
        lng序号 = lng序号 + 1
        i = i + 1
    Loop
    GetPlainTextSaveSQL = True
End Function

'################################################################################################################
'## 功能：  内容的复制操作（包括文本和要素）
'################################################################################################################
Private Sub ExecCopy()
    If Me.edt内容.ReadOnly Then Exit Sub
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bBeteenKeys As Boolean, sKeyType As String, bNeeded As Boolean
    Dim lS As Long, lE As Long, i As Long, bFinded As Boolean, lngLen As Long, lngSum As Long
    
    '扩展起始位置和终止位置，使得其包含完整的要素定义
    lS = Me.edt内容.Selection.StartPos
    lE = Me.edt内容.Selection.EndPos
    bBeteenKeys = IsBetweenAnyKeys(edt内容, lS + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
    If bBeteenKeys Then lS = lKSS
    bBeteenKeys = IsBetweenAnyKeys(edt内容, lE + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
    If bBeteenKeys Then lE = lKEE
    
    '先拷贝RTF内容
    gfrmPublic.edtPublic.NewDoc
    gfrmPublic.edtPublic.ForceEdit = True
    gfrmPublic.edtPublic.TOM.TextDocument.Selection.FormattedText = Me.edt内容.TOM.TextDocument.Range(lS, lE).FormattedText
    '拷贝要素，过滤其他元素（图片、诊断、表格等），关键字也要拷贝过去，保证与内容的隐藏关键字Key值一致！
    Set gfrmPublic.Elements = New cEPRElements
    lngSum = 0
    For i = lS To lE
        bFinded = FindNextAnyKey(edt内容, i + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
        If bFinded Then
            If lKSS < lE Then
                '范围内存在关键字
                If sKeyType = "E" Then
                    '如果是要素，那么拷贝到缓冲区
                    gfrmPublic.Elements.AddExistNode Elements("K" & lKey).Clone(True), True
                Else
                    '如果是其他元素，则清除之（在gfrmPublic.edtPublic中清除，并记录当前位置）！
                    gfrmPublic.edtPublic.Range(lKSS - lS - lngSum, lKEE - lS - lngSum) = ""
                    lngSum = lngSum + lKEE - lKSS   '记录删除内容的总长度
                End If
            Else
                '否则，超出范围，退出循环
                Exit For
            End If
            i = lKEE - 1
        Else
            '不存在任何元素，那么退出循环
            Exit For
        End If
    Next
    Clipboard.Clear
End Sub

'################################################################################################################
'## 功能：  内容的剪切操作（包括文本和要素）
'################################################################################################################
Private Sub ExecCut()
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bBeteenKeys As Boolean, sKeyType As String, bNeeded As Boolean
    Dim lS As Long, lE As Long, i As Long, bFinded As Boolean, lngNum As Long, lngSum As Long
    
    '扩展起始位置和终止位置，使得其包含完整的要素定义
    lS = Me.edt内容.Selection.StartPos
    lE = Me.edt内容.Selection.EndPos
    bBeteenKeys = IsBetweenAnyKeys(edt内容, lS + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
    If bBeteenKeys Then lS = lKSS
    bBeteenKeys = IsBetweenAnyKeys(edt内容, lE + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
    If bBeteenKeys Then lE = lKEE
    
    '先拷贝RTF内容
    gfrmPublic.edtPublic.NewDoc
    gfrmPublic.edtPublic.ForceEdit = True
    gfrmPublic.edtPublic.TOM.TextDocument.Selection.FormattedText = Me.edt内容.TOM.TextDocument.Range(lS, lE).FormattedText
    '拷贝要素，过滤其他元素（图片、诊断、表格等），关键字也要拷贝过去，保证与内容的隐藏关键字Key值一致！
    Set gfrmPublic.Elements = New cEPRElements
    lngSum = 0
    For i = lS To lE
        bFinded = FindNextAnyKey(edt内容, i + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
        If bFinded Then
            If lKSS < lE Then
                '范围内存在关键字
                If sKeyType = "E" Then
                    '如果是要素，那么拷贝到缓冲区
                    gfrmPublic.Elements.AddExistNode Elements("K" & lKey), True
                Else
                    '如果是其他元素，则清除之（在gfrmPublic.edtPublic中清除，并记录当前位置）！
                    gfrmPublic.edtPublic.Range(lKSS - lS - lngSum, lKEE - lS - lngSum) = ""
                    lngSum = lngSum + lKEE - lKSS   '记录删除内容的总长度
                End If
            Else
                '否则，超出范围，退出循环
                Exit For
            End If
            i = lKEE - 1
        Else
            '不存在任何元素，那么退出循环
            Exit For
        End If
    Next
    
    '删除选中内容
    Dim bForce As Boolean, COLOR As OLE_COLOR, bProtect1 As Boolean, bProtect2 As Boolean
    bForce = Me.edt内容.ForceEdit
    Me.edt内容.ForceEdit = True
    Me.edt内容.Range(lS, lE) = ""
    Me.edt内容.ForceEdit = bForce
    Clipboard.Clear
End Sub

'################################################################################################################
'## 功能：  内容的删除操作
'################################################################################################################
Private Sub ExecDelete()
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bBeteenKeys As Boolean, sKeyType As String, bNeeded As Boolean
    Dim lS As Long, lE As Long, i As Long, j As Long, bFinded As Boolean, lngNum As Long, lngSum As Long
    
    '扩展起始位置和终止位置，使得其包含完整的要素定义
    lS = Me.edt内容.Selection.StartPos
    lE = Me.edt内容.Selection.EndPos
    bBeteenKeys = IsBetweenAnyKeys(edt内容, lS + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
    If bBeteenKeys Then lS = lKSS
    bBeteenKeys = IsBetweenAnyKeys(edt内容, lE + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
    If bBeteenKeys Then lE = lKEE
    
    '删除选中内容
    Dim bForce As Boolean, COLOR As OLE_COLOR, bProtect1 As Boolean, bProtect2 As Boolean
    bForce = Me.edt内容.ForceEdit
    Me.edt内容.ForceEdit = True
    
    If Me.edt内容.SelLength > 0 Then
        '选中内容非空
        '扩展起始位置和终止位置，使得其包含完整的要素定义
        lS = Me.edt内容.Selection.StartPos
        lE = Me.edt内容.Selection.EndPos
        bBeteenKeys = IsBetweenAnyKeys(edt内容, lS + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
        If bBeteenKeys Then lS = lKSS
        bBeteenKeys = IsBetweenAnyKeys(edt内容, lE + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
        If bBeteenKeys Then lE = lKEE
        
        Me.edt内容.Freeze
        '非修订模式，则清除所有要素、图片、表格、诊断，不能删除提纲
        lngSum = 0
        For i = lS To lE - 1
            bFinded = FindNextAnyKey(edt内容, i + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
            If bFinded Then
                If lKSS < lE Then   '范围内存在关键字
                    '1、先处理前面的文字
                    Me.edt内容.Range(i, lKSS) = ""
                    lngNum = lKSS - i
                    lE = lE - lngNum
                    lngSum = lngSum + lngNum
                    i = lKSS - lngNum - 1
                    '2、处理后面一个要素、图片、表格、诊断
                    Select Case sKeyType
                    Case "E"    '要素
                        If Elements("K" & lKey).保留对象 = False Then
                            Me.edt内容.Range(lKSS - lngNum, lKEE - lngNum) = ""
                            Elements.Remove "K" & lKey
                            lngSum = lngSum + (lKEE - lKSS)
                            lE = lE - (lKEE - lKSS)
                        Else
                            i = lKEE - lngNum - 1
                        End If
                    Case Else
                       '如果是其他元素，则不处理
                       i = lKEE - lngNum - 1
                    End Select
                Else
                    '否则，超出范围，退出循环
                    Exit For
                End If
            Else
                '不存在任何元素，那么退出循环
                Exit For
            End If
        Next
        If i < lE Then
            Me.edt内容.Range(i, lE) = ""
            lngNum = lE - i
        End If
        Me.edt内容.UnFreeze
        Me.edt内容.SelLength = 0
        Me.edt内容.Range(lE - lngNum, lE - lngNum).Selected
        Clipboard.Clear
    Else
        '没有选择文本
        bBeteenKeys = IsBetweenAnyKeys(edt内容, Me.edt内容.SelStart + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
        If bBeteenKeys Then
            '删除单个诊治要素
            Select Case sKeyType
            Case "E"
                Elements.Remove "K" & lKey
            Case Else
                GoTo LL
            End Select
            Me.edt内容.Range(lKSS, lKEE) = ""
            If Me.edt内容.Range(lKSS - 2, lKSS) = vbCrLf And Me.edt内容.Range(lKSS - 2, lKSS).Font.Protected Then
                Me.edt内容.Range(lKSS - 2, lKSS) = ""
                Me.edt内容.Range(lKSS - 2, lKSS - 2).Font.Protected = False
            Else
                Me.edt内容.Range(lKSS, lKSS).Font.Protected = False
            End If
        Else
            '删除文本
            i = Me.edt内容.SelStart
            j = Len(edt内容.Text)
            
            If Me.edt内容.Range(i, i + 1).Font.Protected = False And (edt内容.Range(i + 1, i + 2).Font.Protected = True Or i = j - 1) Then
                Me.edt内容.Range(i, i + 1) = ""
            ElseIf Me.edt内容.Range(i - 1, i).Font.Protected = True And Me.edt内容.Range(i, i + 1).Font.Protected = False Then
                If Me.edt内容.Range(i, i + 2) = vbCrLf And Me.edt内容.Range(i, i + 2).Font.Protected = False Then
                    Me.edt内容.Range(i, i + 2) = ""
                    Me.edt内容.Range(i, i).Font.Protected = False
                Else
                    Me.edt内容.Delete
                End If
            ElseIf Me.edt内容.Range(i, i + 2) = vbCrLf And Me.edt内容.Range(i, i + 2).Font.Protected = False Then
                Me.edt内容.Range(i, i + 2) = ""
                Me.edt内容.Range(i, i).Font.Protected = False
            ElseIf Me.edt内容.Range(i, i).Font.Protected = False And Me.edt内容.Range(i, i + 1).Font.Protected = False Then
                Me.edt内容.Delete
            ElseIf Me.edt内容.Range(i, i + 2) = vbCrLf And Me.edt内容.Range(i, i + 2).Font.Protected Then
                Me.edt内容.Range(i + 2, i + 2).Selected
            Else
                Me.edt内容.Range(i + 1, i + 1).Selected
            End If
        End If
    End If
LL:
    Me.edt内容.ForceEdit = bForce
End Sub

'################################################################################################################
'## 功能：  内容的粘贴操作（修正要素关键字，对于删除的要素要表现为新增，修订文本也统一改为新增文本）
'################################################################################################################
Private Sub ExecPaste(ByRef edtThis As Object)
    If edtThis.ReadOnly Then Exit Sub
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bBeteenKeys As Boolean, sKeyType As String, bNeeded As Boolean
    Dim i As Long, bForce As Boolean, bFinded As Boolean, strTmp As String, lS As Long, lE As Long, lngLen As Long
    Dim ParaFmt As New cParaFormat
    bBeteenKeys = IsBetweenAnyKeys(edtThis, edtThis.SelStart + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
    If bBeteenKeys Then Exit Sub    '不允许粘贴到元素内部
    
    If edtThis.Selection.Font.ForeColor = tomUndefined Or edtThis.Selection.Font.Protected Then Exit Sub
    
    '如果剪贴板为空，那么就粘贴内部数据
    Dim strClipboard As String
    strClipboard = Clipboard.GetText
    If Len(Trim(strClipboard)) > 0 Then
        '粘贴剪贴板数据
        lS = edtThis.Selection.StartPos
        lE = lS + Len(strClipboard)
        edtThis.ForceEdit = True
        edtThis.Range(lS, edtThis.Selection.EndPos).Text = strClipboard
        edtThis.Range(lS, lE).Font.Strikethrough = False
        edtThis.Range(lS, lE).Font.Protected = False
        edtThis.Range(lS, lE).Font.ForeColor = tomAutoColor
        edtThis.ForceEdit = False
        edtThis.Range(lE, lE).Selected
        Exit Sub
    End If
    
    '先修正关键字
    gfrmPublic.edtPublic.ForceEdit = True
    For i = 1 To gfrmPublic.Elements.Count
        '加入要素
        lKey = Elements.AddExistNode(gfrmPublic.Elements(i).Clone, False)
        Elements("K" & lKey).开始版 = 1
        Elements("K" & lKey).终止版 = 0     '去掉终止版
        Elements("K" & lKey).保留对象 = False
        Elements("K" & lKey).ID = 0
        '修正关键字
        bFinded = FindKey(gfrmPublic.edtPublic, "E", gfrmPublic.Elements(i).Key, lKSS, lKSE, lKES, lKEE, bNeeded)
        If bFinded Then
            strTmp = Format(lKey, "00000000") & "," & IIf(Elements("K" & lKey).保留对象, 1, 0) & ",0)"
            gfrmPublic.edtPublic.Range(lKSS, lKSE) = "ES(" & strTmp
            gfrmPublic.edtPublic.Range(lKES, lKEE) = "EE(" & strTmp
            gfrmPublic.Elements(i).Key = lKey '更新文本的同时，更新Key
        End If
    Next
    
    '拷贝RTF内容，清除前景色和删除线
    bForce = edtThis.ForceEdit
    edtThis.Freeze
    edtThis.ForceEdit = True
    
    lS = 0: lE = Len(gfrmPublic.edtPublic.Text)
    For i = lS To lE - 1
        If gfrmPublic.edtPublic.Range(i, i + 1).Font.ForeColor = PROTECT_FORECOLOR And gfrmPublic.edtPublic.Range(i, i + 1).Font.Protected Then
            '保护文本去掉保护
            gfrmPublic.edtPublic.Range(i, i + 1).Font.Protected = False
        End If
        gfrmPublic.edtPublic.Range(i, i + 1).Font.ForeColor = tomAutoColor
    Next
    
    gfrmPublic.edtPublic.SelectAll
    gfrmPublic.edtPublic.Selection.Font.Strikethrough = False
    lS = edtThis.SelStart

    lngLen = Len(gfrmPublic.edtPublic.Text)
    If lngLen > 0 Then
        edtThis.TOM.TextDocument.Selection.FormattedText = gfrmPublic.edtPublic.TOM.TextDocument.Selection.FormattedText
        '去掉末尾的回车换行符
        If edtThis.Range(lS + lngLen, lS + lngLen + 2) = vbCrLf And edtThis.Range(lS + lngLen, lS + lngLen + 2).Font.Protected = False Then
            edtThis.Range(lS + lngLen, lS + lngLen + 2) = ""
        End If
        edtThis.Range(lS + lngLen, lS + lngLen).Selected
    End If
    lngLen = Len(edt内容.Text)
    Me.edt内容.Range(0, lngLen).Para.SetIndents 0, 0, 0
    Me.edt内容.Range(0, lngLen).Para.SetLineSpacing cprLSSignle, 1
    Me.edt内容.Range(0, lngLen).Para.ListType = cprLTNone
    Me.edt内容.Range(0, lngLen).Font.Size = 10.5
    Me.edt内容.Range(0, lngLen).Font.Name = "宋体"
    
    Me.edt内容.ForceEdit = bForce
    Me.edt内容.UnFreeze
End Sub


'-----------------------------------------------------
'以下为控件事件处理
'-----------------------------------------------------
Private Sub cbo科室_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim sType As String, lSS As Long, lSE As Long, lES As Long, lEE As Long, lKey As Long, bInKeys As Boolean, bNeeded As Boolean
'    On Error Resume Next
    Select Case Control.ID
    Case conMenu_Edit_NewItem
        bInKeys = IsBetweenAnyKeys(Me.edt内容, Me.edt内容.SelStart + 1, sType, lSS, lSE, lES, lEE, lKey, bNeeded)
        If bInKeys Then Exit Sub
        If bInKeys = False Then mfrmInsElement.ShowMe Me, , True, False, True
    Case conMenu_Edit_Modify
        bInKeys = IsBetweenAnyKeys(Me.edt内容, Me.edt内容.SelStart + 1, sType, lSS, lSE, lES, lEE, lKey, bNeeded)
        If bInKeys Then
            mfrmInsElement.Tag = lKey
            mfrmInsElement.ShowMe Me, Elements("K" & lKey), True, False, True
        End If
    Case ID_EDIT_CUT
        ExecCut
    Case ID_EDIT_COPY
        ExecCopy
    Case ID_EDIT_PASTE
        ExecPaste Me.edt内容
    Case ID_INSERT_AUTORECOGNISE                 '智能识别
        '自动识别诊治要素或者字典项目
        Dim strAuto As String
        strAuto = Trim(Me.edt内容.SelText)
        If strAuto = "" Then Exit Sub
        If Len(strAuto) > 100 Then strAuto = Left(strAuto, 100)
        ShowAutoRecSelector strAuto
    Case conMenu_Edit_Delete
        Call ExecDelete
    Case conMenu_Edit_Import
        If Control.Type <> xtpControlComboBox Then Exit Sub
        Call InsertPhrase(Control.ItemData(Control.ListIndex))
    End Select
    Me.edt内容.SetFocus
End Sub

Private Sub cbsThis_Resize()
    Dim lngScaleLeft As Long, lngScaleTop As Long, lngScaleRight As Long, lngScaleBottom As Long
    Call Me.cbsThis.GetClientRect(lngScaleLeft, lngScaleTop, lngScaleRight, lngScaleBottom)
    Me.edt内容.Move lngScaleLeft, lngScaleTop, lngScaleRight - lngScaleLeft, lngScaleBottom - lngScaleTop
    Me.edt内容.PaperWidth = lngScaleRight - lngScaleLeft
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False
    If Me.edt内容.Modified Then
        If MsgBox("词句示范内容已经被修改，是否保存？", vbYesNo + vbQuestion, gstrSysName) = vbYes Then
            Call cmdOK_Click
            Exit Sub
        End If
    End If
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    Dim strText As String, ArraySQL() As String, i As Long, lngCount As Long, blnTran As Boolean
    
    If Trim(Me.txt编号.Text) = "" Then MsgBox "请输入编号！", vbInformation, gstrSysName: Me.txt编号.SetFocus: Exit Sub
'    If Len(Me.txt编号.Text) < Me.txt编号.MaxLength Then MsgBox "编号长度不足！", vbInformation, gstrSysName: Me.txt编号.SetFocus: Exit Sub
    If Trim(Me.txt名称.Text) = "" Then MsgBox "请输入名称！", vbInformation, gstrSysName: Me.txt名称.SetFocus: Exit Sub
    If LenB(StrConv(Trim(Me.txt名称.Text), vbFromUnicode)) > Me.txt名称.MaxLength Then
        MsgBox "名称超长（最多" & Me.txt名称.MaxLength & "个字符或等长的汉字）！", vbInformation, gstrSysName: Me.txt名称.SetFocus: Exit Sub
    End If
    If Me.cbo科室.ListIndex = -1 Then MsgBox "请输入科室！", vbInformation, gstrSysName: Me.cbo科室.SetFocus: Exit Sub
    
    '数据保存
    If Me.Tag = "新增" Then
        mlngWordId = zlDatabase.GetNextId("病历词句示范")
        gstrSQL = mlngWordId & "," & Val(Me.txt分类.Tag) & ",'" & Trim(Me.txt编号.Text) & "','" & Trim(Me.txt名称.Text) & "'"
        If Me.opt范围(0).Value Then
            gstrSQL = gstrSQL & ",0"
        ElseIf Me.opt范围(1).Value Then
            gstrSQL = gstrSQL & ",1"
        Else
            gstrSQL = gstrSQL & ",2"
        End If
        gstrSQL = gstrSQL & "," & Me.cbo科室.ItemData(Me.cbo科室.ListIndex) & "," & Me.lbl人员.Tag
        gstrSQL = "Zl_病历词句示范_Edit(1," & gstrSQL & ")"
    Else
        gstrSQL = mlngWordId & "," & Val(Me.txt分类.Tag) & ",'" & Trim(Me.txt编号.Text) & "','" & Trim(Me.txt名称.Text) & "'"
        If Me.opt范围(0).Value Then
            gstrSQL = gstrSQL & ",0"
        ElseIf Me.opt范围(1).Value Then
            gstrSQL = gstrSQL & ",1"
        Else
            gstrSQL = gstrSQL & ",2"
        End If
        gstrSQL = gstrSQL & "," & Me.cbo科室.ItemData(Me.cbo科室.ListIndex)
        gstrSQL = "Zl_病历词句示范_Edit(2," & gstrSQL & ")"
    End If
    
    '获取SQL语句数组
    ReDim ArraySQL(1 To 2) As String
    ArraySQL(1) = gstrSQL
    
    '前期处理
    ArraySQL(2) = "Zl_病历词句组成_Beforesave(" & mlngWordId & ")"
    
    '获取保存SQL数组
    Call GetSaveSQL(ArraySQL)
    
    '后期处理
    lngCount = UBound(ArraySQL) + 1
    ReDim Preserve ArraySQL(1 To lngCount) As String
    gstrSQL = "Zl_病历词句组成_Aftersave(" & mlngWordId & ")"
    ArraySQL(lngCount) = gstrSQL
    
    '执行保存操作
    Err = 0: On Error GoTo ErrHand
    gcnOracle.BeginTrans
    blnTran = True
    For i = 1 To UBound(ArraySQL)
        gstrSQL = ArraySQL(i)
        Call zlDatabase.ExecuteProcedure(gstrSQL, "cEPRDocument")
    Next
    gcnOracle.CommitTrans
    blnTran = False
    mblnOK = True: Me.Hide
    Exit Sub

ErrHand:
    If blnTran Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub GetSaveSQL(ByRef ArraySQL() As String)
    '获取保存SQL语句
    Dim lKSS As Long, lKSE As Long, lKES As Long, lKEE As Long, lKey As Long, bFinded As Boolean, sKeyType As String, bNeeded As Boolean
    Dim lngEnd As Long, M As Long, N As Long, p As Long
    Dim lng序号 As Long, strText As String
    Dim lngCount As Long
    
    lng序号 = 1     '按照CRLF来分段
    strText = Me.edt内容.Text
    p = 0
    lngEnd = Len(Me.edt内容.Text)
    Do While p < lngEnd
        '获取关键字位置 M
        bFinded = FindNextAnyKey(Me.edt内容, p + 1, sKeyType, lKSS, lKSE, lKES, lKEE, lKey, bNeeded)
        If bFinded Then
            M = lKSS
        Else
            M = lngEnd
        End If
        
        '获取vbCrlf位置 N
        N = InStr(p + 1, strText, vbCrLf, vbTextCompare)
        If N > 0 Then
            N = N - 1
        Else
            N = lngEnd
        End If
        
        If M < N Then
            '保存文本
            Call GetPlainTextSaveSQL(ArraySQL, Me.edt内容.Range(p, M), lng序号)    '序号自动增1
            '保存对象
            If bFinded Then
                Select Case sKeyType
                Case "E"
                    Elements("K" & lKey).对象序号 = lng序号
                    p = lKEE    '调整当前位置
                    With Elements("K" & lKey)
                        lngCount = UBound(ArraySQL) + 1
                        ReDim Preserve ArraySQL(1 To lngCount) As String
                        gstrSQL = "Zl_病历词句组成_Insert(" & mlngWordId & "," & lng序号 & ",1,'" & .内容文本 & "','" & .要素名称 & "'," & _
                            IIf(.诊治要素ID = 0, "NULL", .诊治要素ID) & "," & .替换域 & "," & .要素类型 & "," & .要素长度 & "," & .要素小数 & ",'" & .要素单位 & "'," & _
                            .要素表示 & ",'" & .要素值域 & "'," & .输入形态 & ",'" & .对象属性 & "')"
                        ArraySQL(lngCount) = gstrSQL
                    End With
                    lng序号 = lng序号 + 1
                End Select
            Else
                p = M
            End If
        Else
            If Me.edt内容.Range(N, N + 2) = vbCrLf And Me.edt内容.Range(N, N + 2).Font.Protected = True Then
                '该回车属于下一个对象（图片或者表格）
                If p < N Then Call GetPlainTextSaveSQL(ArraySQL, Me.edt内容.Range(p, N), lng序号)                 '序号自动增1
            Else
                '保存文本
                Call GetPlainTextSaveSQL(ArraySQL, Me.edt内容.Range(p, IIf(N >= lngEnd, N, N + 2)), lng序号) '序号自动增1
            End If
            p = N + 2
        End If
    Loop
End Sub

Private Sub cmd分类_Click()
    With Me.tvw分类
        .Left = Me.txt分类.Left: .Width = Me.txt分类.Width
        .Top = Me.txt分类.Top + Me.txt分类.Height: .Height = Me.pic内容.Height
        .ZOrder 0: .Visible = True: .SetFocus
    End With
End Sub

Private Sub edt内容_Change(ViewMode As zlRichEditor.ViewModeEnum)
    If mlngHP > 0 Then
        If blnSpaceEvent Then
            blnSpaceEvent = False
            Exit Sub
        Else
            '恢复空格！
            If Me.edt内容.Range(mlngHP - 1, mlngHP).Font.Hidden And _
                Me.edt内容.Range(mlngHP, mlngHP + 1).Font.Hidden = False And _
                Me.edt内容.Range(mlngHP, mlngHP + 1) = " " Then
                
                Dim blnForce As Boolean
                blnForce = Me.edt内容.ForceEdit
                Me.edt内容.ForceEdit = True
                Me.edt内容.Range(mlngHP, mlngHP + 1) = ""
                Me.edt内容.Range(mlngHP, mlngHP).Font.Protected = False
                Me.edt内容.Range(mlngHP, mlngHP).Font.Hidden = False
                Me.edt内容.ForceEdit = blnForce
            End If
            mlngHP = 0
            blnSpaceEvent = False
        End If
    End If
End Sub

Private Sub edt内容_KeyDown(ViewMode As zlRichEditor.ViewModeEnum, KeyCode As Integer, Shift As Integer)
    If Me.edt内容.SelLength > 0 Then Exit Sub
    If Shift <> 0 Then Exit Sub
    Select Case KeyCode
    Case 0, vbKeyUp, vbKeyDown, vbKeyLeft, vbKeyRight, vbKeyHome, vbKeyEnd, _
        vbKeyDelete, vbKeyBack, vbKeyTab, vbKeyInsert, vbKeyPageDown, vbKeyPageUp, _
        vbKeyPause, vbKeyPrint, vbKeyNumlock, vbKeyScrollLock, vbKeyCapital, _
        vbKeyF1, vbKeyF2, vbKeyF3, vbKeyF4, vbKeyF5, vbKeyF6, vbKeyF7, vbKeyF8, vbKeyF9, vbKeyF10, vbKeyF11, vbKeyF12
        
        DoEvents
        Exit Sub
    End Select
    
    '发现在隐藏关键字后面，则自动加一个空格（非保护和隐藏属性）
    Dim i As Long, blnForce As Boolean
    With Me.edt内容
        blnForce = .ForceEdit
        i = .SelStart
    
LL1:
        If .Range(i - 1, i).Font.Hidden And _
            .Range(i, i + 1).Font.Hidden = False And _
            .Range(i, i + 1).Font.Protected = False Then
            'A问题：（隐藏文本）|普通文本
            
            mlngHP = i
            .ForceEdit = True
            .Range(i, i).Font.Protected = False
            .Range(i, i).Font.Hidden = False
            blnSpaceEvent = True
            .Range(i, i) = " "
            .Range(i + 1, i + 1).Selected
            .ForceEdit = blnForce
        Else
            If .Range(i - 1, i).Font.Hidden And _
                .Range(i, i + 1).Font.Hidden = False And _
                .Range(i, i + 1).Font.Protected Then
                'B问题1：普通文本（隐藏文本）|（保护文本）（隐藏文本）普通文本
                i = i - 16
                If .Range(i - 1, i + 3) Like ")?S(" And _
                    .Range(i - 1, i + 3).Font.Hidden = True Then
                    'C问题：（隐藏文本）（保护文本）（隐藏文本）|（隐藏文本）（保护文本）（隐藏文本）
                    mlngHP = -1
                    .ForceEdit = True
                    .Range(i, i).Font.Protected = False
                    .Range(i, i).Font.Hidden = False
                    .Range(i - 1, i).Font.ForeColor = vbBlack
                    blnSpaceEvent = True
                    .Range(i, i) = IIf(.Range(i - 16, i - 14) <> "EE", " ", "，")
                    .Range(i, i + 1).Font.Protected = False
                    .Range(i, i + 1).Font.Hidden = False
                    .Range(i, i + 1).Font.ForeColor = vbBlack
                    .Range(i + 1, i + 1).Selected
                    .ForceEdit = blnForce
                ElseIf .Range(i + 1, i + 3) = "E(" And .Range(i, i + 3).Font.Protected And _
                    .Range(i + 16, i + 18) = vbCrLf And .Range(i + 16, i + 18).Font.Protected Then
                    'D问题：提纲后面跟图片，在之间没有文字时，无法插入其他文字
                    i = i + 16
                    mlngHP = -1
                    .ForceEdit = True
                    .Range(i, i).Font.Protected = False
                    .Range(i, i).Font.Hidden = False
                    .Range(i - 1, i).Font.ForeColor = vbBlack
                    blnSpaceEvent = True
                    .Range(i, i) = IIf(.Range(i - 16, i - 14) <> "EE", " ", "，")
                    .Range(i, i + 1).Font.Protected = False
                    .Range(i, i + 1).Font.Hidden = False
                    .Range(i, i + 1).Font.ForeColor = vbBlack
                    If (.Range(i - 16, i - 14) <> "EE") Then
                        .Range(i, i + 1).Selected
                    Else
                        .Range(i + 1, i + 1).Selected
                    End If
                    .ForceEdit = blnForce
                Else
                    .Range(i, i).Selected
                End If
            ElseIf .Range(i - 1, i).Font.Hidden = False And _
                .Range(i - 1, i).Font.Protected And _
                .Range(i, i + 1).Font.Hidden Then
                'B问题2：普通文本（隐藏文本）（保护文本）|（隐藏文本）普通文本
                i = i + 16
                If .Range(i - 1, i + 3) Like ")?S(" And _
                    .Range(i - 1, i + 3).Font.Hidden = True Then
                    'C问题：（隐藏文本）（保护文本）（隐藏文本）|（隐藏文本）（保护文本）（隐藏文本）
                    mlngHP = -1
                    .ForceEdit = True
                    .Range(i, i).Font.Protected = False
                    .Range(i, i).Font.Hidden = False
                    .Range(i - 1, i).Font.ForeColor = vbBlack
                    blnSpaceEvent = True
                    .Range(i, i) = IIf(.Range(i - 16, i - 14) <> "EE", " ", "，")
                    .Range(i, i + 1).Font.Protected = False
                    .Range(i, i + 1).Font.Hidden = False
                    .Range(i, i + 1).Font.ForeColor = vbBlack
                    .Range(i + 1, i + 1).Selected
                    .ForceEdit = blnForce
                Else
                    GoTo LL1
                End If
            ElseIf .Range(i - 1, i).Font.Hidden = False And .Range(i, i + 2) = vbCrLf And .Range(i, i + 2).Font.Protected Then
                mlngHP = -1
                .ForceEdit = True
                .Range(i, i).Font.Protected = False
                .Range(i, i).Font.Hidden = False
                .Range(i - 1, i).Font.ForeColor = vbBlack
                blnSpaceEvent = True
                .Range(i, i) = IIf(.Range(i - 16, i - 14) <> "EE", " ", "，")
                .Range(i, i + 1).Font.Protected = False
                .Range(i, i + 1).Font.Hidden = False
                .Range(i, i + 1).Font.ForeColor = vbBlack
                If (.Range(i - 16, i - 14) <> "EE") Then
                    .Range(i, i + 1).Selected
                Else
                    .Range(i + 1, i + 1).Selected
                End If
                .ForceEdit = blnForce
            End If
        End If
    End With

End Sub

Private Sub edt内容_RequestRightMenu(ViewMode As zlRichEditor.ViewModeEnum, Shift As Integer, x As Single, y As Single)
    Dim Popup As CommandBar, Control As CommandBarControl
    Set Popup = cbsThis.Add("Popup", xtpBarPopup)
    With Popup.Controls
        Set Control = .Add(xtpControlButton, ID_EDIT_CUT, "剪切(&X)")
        Set Control = .Add(xtpControlButton, ID_EDIT_COPY, "复制(&C)")
        Set Control = .Add(xtpControlButton, ID_EDIT_PASTE, "粘贴(&V)    ")
        Popup.ShowPopup
    End With
End Sub

Private Sub Form_Activate()
    If blnActive = False Then
        If Me.txt编号.Visible And Me.txt编号.Enabled Then Me.txt编号.SetFocus
        blnActive = True
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyEscape Then Exit Sub
    If Me.tvw分类.Visible Then
        Me.tvw分类.Visible = False: Me.txt分类.SetFocus: Exit Sub
    End If
    Call cmdCancel_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Me.edt内容.PaperWidth = 16840
    Me.edt内容.ResetWYSIWYG
    Set mfrmInsElement = New frmInsElement
    Set Elements = New cEPRElements
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    If Me.Width < Me.cmdOK.Left + Me.cmdOK.Width Then Me.Width = Me.cmdOK.Left + Me.cmdOK.Width
    If Me.Height < Me.pic内容.Top + 2000 Then Me.Height = Me.pic内容.Top + 2000
    With Me.pic内容
        .Left = 0: .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - .Top
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Elements = Nothing
    Unload mfrmInsElement
    Set mfrmInsElement = Nothing
End Sub

Private Sub mfrmInsElement_pCancel()
    mfrmInsElement.Hide
    mfrmInsElement.Tag = ""
End Sub

Private Sub mfrmInsElement_pOK(Ele As cEPRElement)
    '插入诊治要素
    Dim lngKey As Long
    If mfrmInsElement.Tag <> "" Then
        '修改模式
        Dim sType As String, lSS As Long, lSE As Long, lES As Long, lEE As Long, lKey As Long, bInKeys As Boolean, bNeeded As Boolean
        bInKeys = FindKey(Me.edt内容, "E", mfrmInsElement.Tag, lSS, lSE, lES, lEE, bNeeded)
        If bInKeys Then
            Elements.Remove "K" & mfrmInsElement.Tag
            With Me.edt内容
                .ForceEdit = True
                .Range(lSS, lEE) = ""
                .Range(lSS, lSS).Font.Protected = False
                .Range(lSS, lSS).Selected
                .ForceEdit = False
            End With
        End If
        lngKey = Elements.AddExistNode(Ele, True)
        Elements("K" & lngKey).InsertIntoEditor Me.edt内容, , False
        bInKeys = FindKey(Me.edt内容, "E", lngKey, lSS, lSE, lES, lEE, bNeeded)
        If bInKeys Then
            If Elements("K" & lngKey).输入形态 = 0 Then
                Me.edt内容.Range(lSE, lES).Selected
            Else
                Me.edt内容.Range(lSE + 1, lSE + 1).Selected
            End If
        End If
    Else
        lngKey = Elements.AddExistNode(Ele)
        Elements("K" & lngKey).InsertIntoEditor Me.edt内容, , True
    End If
    mfrmInsElement.Tag = ""
End Sub

Private Sub opt范围_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub tvw分类_DblClick()
    If Me.tvw分类.SelectedItem Is Nothing Then Exit Sub
    Me.txt分类.Tag = Mid(Me.tvw分类.SelectedItem.Key, 2)
    Me.txt分类.Text = Me.tvw分类.SelectedItem.Text
    Me.txt分类.SetFocus
    Call zlDefaultCode
End Sub

Private Sub tvw分类_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn
        If Me.tvw分类.SelectedItem Is Nothing Then Exit Sub
        If Me.tvw分类.SelectedItem.Children > 0 Then Exit Sub
        Call tvw分类_DblClick
    Case vbKeySpace
        Call tvw分类_DblClick
    Case vbKeyEscape
        Call tvw分类_LostFocus
    End Select
End Sub

Private Sub tvw分类_LostFocus()
    If Me.cmd分类 Is ActiveControl Then Exit Sub
    Me.tvw分类.Visible = False
End Sub

Private Sub txt编号_Change()
    ValidControlText txt编号
End Sub

Private Sub txt编号_GotFocus()
    Me.txt编号.SelStart = 0: Me.txt编号.SelLength = 100
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt编号_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt分类_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txt名称_Change()
    ValidControlText txt名称
End Sub

Private Sub txt名称_GotFocus()
    Me.txt名称.SelStart = 0: Me.txt名称.SelLength = 4000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt名称_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Trim(Me.edt内容.Text) = "" Then Me.edt内容.Text = Trim(Me.txt名称.Text)
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    End If
    If InStr("%_'""", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

