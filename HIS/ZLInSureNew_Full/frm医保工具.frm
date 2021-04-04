VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm医保工具 
   Caption         =   "医保工具"
   ClientHeight    =   6075
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8940
   Icon            =   "frm医保工具.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6075
   ScaleWidth      =   8940
   Begin VB.PictureBox picSplitV 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5490
      Left            =   2490
      MousePointer    =   9  'Size W E
      ScaleHeight     =   5490
      ScaleWidth      =   45
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   45
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3360
      Top             =   1950
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保工具.frx":1CFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保工具.frx":3A04
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保工具.frx":3C1E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwMain 
      Height          =   5715
      Left            =   2550
      TabIndex        =   0
      Top             =   0
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   10081
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      Icons           =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList ils32 
      Left            =   690
      Top             =   4380
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保工具.frx":3E38
            Key             =   "Fix"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保工具.frx":4152
            Key             =   "FixD"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保工具.frx":446C
            Key             =   "Common"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保工具.frx":4786
            Key             =   "CommonD"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保工具.frx":4AA0
            Key             =   "Disease"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   90
      Top             =   4380
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
            Picture         =   "frm医保工具.frx":537A
            Key             =   "Fix"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保工具.frx":5694
            Key             =   "FixD"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保工具.frx":59AE
            Key             =   "Common"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保工具.frx":5CC8
            Key             =   "CommonD"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保工具.frx":5FE2
            Key             =   "Disease"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm医保工具.frx":657C
            Key             =   "Limit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwKind_S 
      Height          =   5715
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2385
      _ExtentX        =   4207
      _ExtentY        =   10081
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ils32"
      SmallIcons      =   "ils16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   5715
      Width           =   8940
      _ExtentX        =   15769
      _ExtentY        =   635
      SimpleText      =   $"frm医保工具.frx":69CE
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frm医保工具.frx":6A15
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10689
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
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
End
Attribute VB_Name = "frm医保工具"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mintInsure As Integer
Private mstrPrivs As String
Private Const clng基础数据 As Long = 1
Private Const clng门诊业务 As Long = 2
Private Const clng住院业务 As Long = 3
Private Const clng其它 As Long = 4
Private Const clng返回 As Long = 9998
Private Const clng退出 As Long = 9000

Private mclsInsure As New clsInsure
Private rsMenu As New ADODB.Recordset       '列：上级ID,ID,标题,权限
'本模块权限清单:上传，下载，基础数据设置，基础数据查询，门诊日结，门诊结算，门诊数据查询，入院，出院，结算，住院数据查询，其它

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    Dim strIcon As String, lst As ListItem
    On Error Resume Next
    
    Call RestoreWinState(Me, App.ProductName)
    
    mstrPrivs = gstrPrivs
    
    '装入保险大类
    gstrSQL = "select 序号,名称,是否固定 from 保险类别 where nvl(是否禁止,0)<>1 And 医保部件 Is NULL order by 序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
    If rsTemp.RecordCount = 0 Then
        '如果是在窗体初始化时调用，就不用处理其它内容了
        MsgBox "没有可用保险类别，不能使用本功能。", vbInformation, gstrSysName
        Unload Me
        Exit Sub
    End If
    
    lvwKind_S.ListItems.Clear
    Do Until rsTemp.EOF
        strIcon = IIf(rsTemp("是否固定") = 1, "Fix", "Common")
        Set lst = lvwKind_S.ListItems.Add(, "K" & rsTemp("序号"), rsTemp("名称"), strIcon, strIcon)
        
        rsTemp.MoveNext
    Loop
    If lvwKind_S.SelectedItem Is Nothing Then lvwKind_S.ListItems(1).Selected = True
    Call lvwKind_S_ItemClick(lvwKind_S.ListItems(1))
End Sub

Private Sub Form_Resize()
    Dim sngTop As Single, sngBottom As Single
    On Error Resume Next
    If Me.WindowState = 1 Then Exit Sub
    
    sngTop = 0
    sngBottom = ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0)
    
    lvwKind_S.Top = sngTop
    lvwKind_S.Height = IIf(sngBottom - lvwKind_S.Top > 0, sngBottom - lvwKind_S.Top, 0)
    lvwKind_S.Left = ScaleLeft
    
    picSplitV.Top = sngTop
    picSplitV.Height = IIf(sngBottom - picSplitV.Top > 0, sngBottom - picSplitV.Top, 0)
    picSplitV.Left = lvwKind_S.Left + lvwKind_S.Width
    
    lvwMain.Top = sngTop
    lvwMain.Left = picSplitV.Left + 35
    lvwMain.Width = ScaleWidth - lvwMain.Left
    lvwMain.Height = picSplitV.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub lvwKind_S_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Call LoadTools
End Sub

Private Sub lvwMain_DblClick()
    Dim lngKey As Long, lngKey_上级 As Long
    Dim str权限 As String           '为空表示不进行权限控制
    Dim blnOwner As Boolean
    Dim lvwItem As ListItem
    
    If lvwMain.ListItems.Count = 0 Then Exit Sub
    If lvwMain.SelectedItem Is Nothing Then Exit Sub
    
    lngKey = CLng(Mid(lvwMain.SelectedItem.Key, 3))
    lngKey_上级 = Val(lvwMain.SelectedItem.Tag)
    
    If lngKey = clng退出 Then
        Unload Me
        Exit Sub
    End If
    
    If lngKey = clng返回 Then
        '显示上一级的数据并退出
        If lngKey_上级 = 0 Then
            lngKey = 0
        Else
            rsMenu.Filter = "ID=" & lngKey_上级
            lngKey = rsMenu!上级ID
            rsMenu.Filter = 0
        End If
    End If
    
    '装入数据
    rsMenu.Filter = "上级ID=" & lngKey
    If rsMenu.RecordCount = 0 Then
        rsMenu.Filter = 0
        Call ExecuteFuncs(lngKey)
        Exit Sub
    End If
    
    lvwMain.ListItems.Clear
    With rsMenu
        Do While Not .EOF
            '判断是否拥有权限
            blnOwner = True
            str权限 = Nvl(rsMenu!权限)
            If str权限 <> "" Then
                blnOwner = (InStr(1, ";" & mstrPrivs & ";", ";" & str权限 & ";") <> 0)
            End If
            
            If blnOwner Then
                Set lvwItem = lvwMain.ListItems.Add(, "K_" & rsMenu!ID, rsMenu!标题, 1)
                lvwItem.Tag = Nvl(rsMenu!上级ID, 0)
            End If
            
            .MoveNext
        Loop
        .MoveFirst
        '加入常项，如退出
        Set lvwItem = lvwMain.ListItems.Add(, "K_" & clng返回, "返回", 2)
        lvwItem.Tag = Nvl(rsMenu!上级ID, 0)
        Set lvwItem = lvwMain.ListItems.Add(, "K_" & clng退出, "退出", 3)
        lvwItem.Tag = Nvl(rsMenu!上级ID, 0)
        If lngKey = 0 Then
            lvwMain.ListItems("K_" & clng返回).Ghosted = True
        Else
            lvwMain.ListItems("K_" & clng返回).Ghosted = False
        End If
        .Filter = 0
    End With
End Sub

Public Sub ShowForm(ByVal frmParent As Object, ByVal intinsure As Integer)
    On Error Resume Next
    mintInsure = intinsure
    Me.Show , frmParent
End Sub

Public Sub InitInsure(ByVal intinsure As Integer)
    mintInsure = intinsure
End Sub

Public Sub LoadTools()
    Dim lngCounts As Long '记录工具数
    If lvwKind_S.ListItems.Count = 0 Then Exit Sub
    If lvwKind_S.SelectedItem Is Nothing Then Exit Sub
    
    mintInsure = Mid(lvwKind_S.SelectedItem.Key, 2)
    lvwMain.ListItems.Clear

    '初始化基础记录
    Call Record_Init(rsMenu, "上级ID," & adDouble & ",18|ID," & adDouble & ",18|标题," & adLongVarChar & ",100|权限," & adLongVarChar & ",100")
    
    lvwMain.ListItems.Add , "K_" & clng基础数据, "基础数据", 1
    Call Record_Add(rsMenu, "上级ID|ID|标题", "0|" & clng基础数据 & "|" & "基础数据")
    lvwMain.ListItems.Add , "K_" & clng门诊业务, "门诊业务", 1
    Call Record_Add(rsMenu, "上级ID|ID|标题", "0|" & clng门诊业务 & "|" & "门诊业务")
    lvwMain.ListItems.Add , "K_" & clng住院业务, "住院业务", 1
    Call Record_Add(rsMenu, "上级ID|ID|标题", "0|" & clng住院业务 & "|" & "住院业务")
    lvwMain.ListItems.Add , "K_" & clng其它, "其它", 1
    Call Record_Add(rsMenu, "上级ID|ID|标题", "0|" & clng其它 & "|" & "其它")
    
    lvwMain.ListItems.Add , "K_" & clng返回, "返回", 2
    lvwMain.ListItems.Add , "K_" & clng退出, "退出", 3
    
    '第一层不允许执行退到上一层的功能
    lvwMain.ListItems("K_" & clng返回).Ghosted = True
    Select Case mintInsure
    Case TYPE_昭通
        Call Record_Add(rsMenu, "上级ID|ID|标题|权限", clng基础数据 & "|100|" & "下载服务项目|下载")
        Call Record_Add(rsMenu, "上级ID|ID|标题|权限", clng基础数据 & "|101|" & "下载材料项目|下载")
        Call Record_Add(rsMenu, "上级ID|ID|标题|权限", clng基础数据 & "|102|" & "下载标准药品目录|下载")
        Call Record_Add(rsMenu, "上级ID|ID|标题|权限", clng基础数据 & "|103|" & "下载药品执行库|下载")
        Call Record_Add(rsMenu, "上级ID|ID|标题|权限", clng门诊业务 & "|200|" & "门诊结帐|门诊日结")
        Call Record_Add(rsMenu, "上级ID|ID|标题|权限", clng门诊业务 & "|201|" & "购药明细查询|")
        Call Record_Add(rsMenu, "上级ID|ID|标题|权限", clng门诊业务 & "|202|" & "手工冲销门诊结算|门诊结算")
        Call Record_Add(rsMenu, "上级ID|ID|标题|权限", clng住院业务 & "|300|" & "住院情况查询(含住院记录,费用明细,费用汇总)|")
        Call Record_Add(rsMenu, "上级ID|ID|标题|权限", clng住院业务 & "|301|" & "医保出院撤销|出院")
        Call Record_Add(rsMenu, "上级ID|ID|标题|权限", clng其它 & "|400|" & "操作员管理|基础数据设置")
        Call Record_Add(rsMenu, "上级ID|ID|标题|权限", clng其它 & "|401|" & "申报药品|基础数据设置")
        Call Record_Add(rsMenu, "上级ID|ID|标题|权限", clng其它 & "|402|" & "操作员复位|基础数据设置")
    Case TYPE_涪陵
        Call Record_Add(rsMenu, "上级ID|ID|标题", clng门诊业务 & "|200|" & "手工冲销门诊结算")
    Case TYPE_铜山县
        Call Record_Add(rsMenu, "上级ID|ID|标题|权限", clng住院业务 & "|300|" & "清除上传标志|")
    Case TYPE_北京尚洋
        Call 医保初始化_北京尚洋
        Call Record_Add(rsMenu, "上级ID|ID|标题|权限", clng基础数据 & "|100|" & "病案数据上传|上传")
        Call Record_Add(rsMenu, "上级ID|ID|标题|权限", clng基础数据 & "|101|" & "下载医保就诊目录|下载")
'        Call Record_Add(rsMenu, "上级ID|ID|标题|权限", clng基础数据 & "|102|" & "价目目录维护|")
    Case TYPE_贵阳市
        Call Record_Add(rsMenu, "上级ID|ID|标题|权限", clng基础数据 & "|100|" & "清算单管理|上传")
        Call Record_Add(rsMenu, "上级ID|ID|标题|权限", clng门诊业务 & "|200|" & "门诊转医保|门诊结算")
        Call Record_Add(rsMenu, "上级ID|ID|标题|权限", clng门诊业务 & "|201|" & "门诊结算数据核对|门诊结算")
        Call Record_Add(rsMenu, "上级ID|ID|标题|权限", clng门诊业务 & "|202|" & "医保黑名单管理|门诊结算")
        Call Record_Add(rsMenu, "上级ID|ID|标题|权限", clng住院业务 & "|300|" & "特殊药品审批|住院数据查询")
        Call Record_Add(rsMenu, "上级ID|ID|标题|权限", clng住院业务 & "|301|" & "结算上传管理|上传")
        Call Record_Add(rsMenu, "上级ID|ID|标题|权限", clng其它 & "|400|" & "冲销本地未成功的中心结算费用|基础数据设置")
        Call Record_Add(rsMenu, "上级ID|ID|标题|权限", clng其它 & "|401|" & "门诊限量用药审批|基础数据设置")
    End Select
End Sub

Private Sub picSplitV_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    
    With picSplitV
        .Move .Left + x
    End With
    Me.lvwKind_S.Width = picSplitV.Left
    Call Form_Resize
End Sub

Private Function ExecuteFuncs(ByVal lng功能号 As Long) As Boolean
    Dim blnDo As Boolean
    On Error GoTo errHand
    
    If lng功能号 < 100 Then Exit Function
    Select Case mintInsure
    Case TYPE_昭通
        blnDo = 昭通工具包(lng功能号)
    Case TYPE_涪陵
        blnDo = 涪陵工具包(lng功能号)
    Case TYPE_铜山县
        blnDo = 铜山县医保工具包(lng功能号)
    Case TYPE_北京尚洋
        blnDo = 尚洋工具包(lng功能号)
    Case TYPE_贵阳市
        blnDo = 贵阳工具包(lng功能号)
    End Select
    
    ExecuteFuncs = blnDo
    If blnDo Then MsgBox "执行成功！", vbInformation, gstrSysName
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function 贵阳工具包(ByVal lng功能号 As Long) As Boolean
    Dim lngID As Long
    Dim str编码 As String, str名称 As String
    Dim nodRowset As MSXML2.IXMLDOMElement, nodRow As MSXML2.IXMLDOMElement
    On Error GoTo errHand
    
    If Not gclsInsure.InitInsure(gcnOracle, TYPE_贵阳市) Then Exit Function
    
    Select Case lng功能号
        Case 100
            Call frm清算单管理_贵阳.ShowME(mintInsure)
        Case 200
            Call frm门诊转医保.ShowME(mintInsure)
        Case 201
            Call frm门诊结算数据核对_贵阳.ShowME(1, mintInsure)
        Case 202
            With frmIdentify贵阳黑名单
                .intinsure = mintInsure
                .Show vbModal
            End With
            Set frmIdentify贵阳黑名单 = Nothing
        Case 101
            '保存到我们的病种目录表中
            If InitXML = False Then Exit Function
            If Not CommServer("QUERYSPECILLNESS") Then Exit Function
            Set nodRowset = mdomOutput.documentElement.selectSingleNode("ROWSET")
            If nodRowset Is Nothing Then Exit Function
            
            '根据编码得到险种名称
            For Each nodRow In nodRowset.childNodes
                lngID = zlDatabase.GetNextID("保险病种")
                str编码 = GetAttributeValue(nodRow, "SPECILLNESSCODE")
                str名称 = GetAttributeValue(nodRow, "SPECILLNESSNAME")
                gstrSQL = "zl_保险病种_INSERT(" & lngID & "," & TYPE_贵阳市 & ",'" & str编码 & "','" & str名称 & "','" & zlCommFun.SpellCode(str名称) & "',2,0,0)"
                Call zlDatabase.ExecuteProcedure(gstrSQL, "更新病种")
            Next
        Case 300
            Call frm特殊项目审批_贵阳.ShowSelect(mintInsure)
        Case 301
            With frm保险结账上传_贵阳
                .Insure = TYPE_贵阳市
                .Show vbModal
            End With
            Set frm保险结账上传_贵阳 = Nothing
        Case 400
            With frmIdentify贵阳补充结算
                .Insure = mintInsure
                .Show vbModal
            End With
            Set frmIdentify贵阳补充结算 = Nothing
        Case 401
            frm限量药品审批_贵阳.ShowME mintInsure
    End Select
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function 尚洋工具包(ByVal lng功能号 As Long) As Boolean
    If Not gclsInsure.InitInsure(gcnOracle, TYPE_北京尚洋) Then Exit Function
    
    Select Case lng功能号
        Case 100    '病案数据上传
            frmMain_北京尚洋病案信息.Show 1, Me
        Case 101
            With frmMain_北京尚洋目录下发
                .intinsure = TYPE_北京尚洋
                .Show 1, Me
            End With
        Case 102
            With frmMain_北京尚洋价目管理
                .Show 1, Me
            End With
    End Select
End Function

Private Function 铜山县医保工具包(ByVal lng功能号 As Long) As Boolean
    Const c铜山县_清除上传标志 As Integer = 300
    Dim rsTmp As New ADODB.Recordset
    Dim lng病人ID As Long, lng主页ID As Long
    Select Case lng功能号
    Case c铜山县_清除上传标志
    
        gstrSQL = "Select a.病人id||'_'||a.主页id as ID,a.病人id, a.主页id, b.住院号, b.住院次数, b.姓名, b.性别, b.年龄, b.身份证号, a.入院日期" & vbNewLine & _
                "From 病案主页 a, 病人信息 b" & vbNewLine & _
                "Where a.病人id = b.病人id And a.主页id = Nvl(b.住院次数, 0) And b.在院=1 And A.险类=" & TYPE_铜山县
        Set rsTmp = zlDatabase.ShowSelect(Me, gstrSQL, 0, "选择病人", True)
        If rsTmp Is Nothing Then
            MsgBox "无在院医保病人可供选择。", vbQuestion, gstrSysName
        Else
            If rsTmp.State = 0 Then
                MsgBox "无在院医保病人可供选择。", vbQuestion, gstrSysName
            Else
                If rsTmp.RecordCount > 0 Then
                    lng病人ID = Nvl(rsTmp.Fields("病人ID"), 0)
                    lng主页ID = Nvl(rsTmp.Fields("主页ID"), 0)
                    If MsgBox("将要清除[" & Nvl(rsTmp.Fields("姓名")) & "]，住院号为（" & Nvl(rsTmp.Fields("住院号")) & "）的费用上传标志，请确定！", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                        gstrSQL = "Update 病人费用记录 Set 是否上传=0 Where 病人ID=" & lng病人ID & " And 主页ID=" & lng主页ID & " & And 是否上传=1"
                        gcnOracle.Execute gstrSQL
                    End If
                End If
            End If
        End If
    End Select
End Function

Private Function 昭通工具包(ByVal lng功能号 As Long) As Boolean
    Const C昭通_下载服务项目 As Integer = 100
    Const C昭通_下载材料项目 As Integer = 101
    Const C昭通_下载标准药品目录 As Integer = 102
    Const C昭通_下载药品执行库 As Integer = 103
    Const C昭通_门诊结帐 As Integer = 200
    Const C昭通_购药明细 As Integer = 201
    Const C昭通_手工冲销门诊单据 As Integer = 202
    Const C昭通_住院情况 As Integer = 300
    Const C昭通_医保出院撤销 As Integer = 301
    Const C昭通_操作员管理 As Integer = 400
    Const C昭通_申报药品 As Integer = 401
    Const C昭通_操作员复位 As Integer = 402
    
    If lng功能号 <> C昭通_操作员管理 Then
        If Not mclsInsure.InitInsure(gcnOracle, TYPE_昭通) Then Exit Function
    Else
        If Not 医保初始化_昭通(False) Then Exit Function
    End If
    
    Select Case lng功能号
    Case C昭通_门诊结帐
        昭通工具包 = frmConn昭通.Execute("I250", 0, "", "正在进行医保门诊结帐......")
        Call ShowWindow(frmConn昭通.hwnd, 0)
    Case C昭通_购药明细
        昭通工具包 = True
        Call frm昭通查询窗体.ShowForm("购药明细查询", "门诊流水号")
    Case C昭通_住院情况
        昭通工具包 = True
        Call frm昭通查询窗体.ShowForm("住院情况查询", "住院流水号")
    Case C昭通_下载服务项目
        昭通工具包 = 昭通_下载服务项目
    Case C昭通_下载材料项目
        昭通工具包 = 昭通_下载材料项目
    Case C昭通_下载标准药品目录
        昭通工具包 = 昭通_下载标准药品目录
    Case C昭通_下载药品执行库
        昭通工具包 = 昭通_下载药品执行库
    Case C昭通_操作员管理
        '操作员管理不必进行医保初始化
        昭通工具包 = 昭通_操作员管理
    Case C昭通_操作员复位
        昭通工具包 = 昭通_操作员复位
    Case C昭通_手工冲销门诊单据
        昭通工具包 = 昭通_手工冲销门诊单据
    Case C昭通_申报药品
        昭通工具包 = 昭通_申报药品
    Case C昭通_医保出院撤销
        Dim StrInput As String
        Dim rsTemp As New ADODB.Recordset
        On Error GoTo errHand
        StrInput = InputBox("请输入该病人的HIS住院号：", "医保出院撤销")
        If Trim(StrInput) = "" Then Exit Function
        
        gstrSQL = " Select A.顺序号 From 保险帐户 A,病人信息 B Where A.病人ID=B.病人ID And B.住院号=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取医保住院号", StrInput)
        If rsTemp.RecordCount = 0 Then Exit Function
        If frmConn昭通.Execute("I345", 0, Nvl(rsTemp!顺序号), "正在进行医保出院撤销......") = False Then Exit Function
        MsgBox "医保出院撤销成功，可以重新办理出院结算业务了！", vbInformation, gstrSysName
        Exit Function
errHand:
        If ErrCenter = 1 Then
            Resume
        End If
        Exit Function
    End Select
End Function

Private Function 昭通_操作员管理() As Boolean
    frmUserList.Show vbModal
End Function

Private Function 昭通_操作员复位() As Boolean
    Dim strCode As String
    strCode = InputBox("请输入待复位操作员编号：", "操作员复位", "00")
    If Trim(strCode) = "" Then Exit Function
    If Len(strCode) > 2 Then Exit Function
    
    If Not frmConn昭通.Execute("I050", 3, strCode, "操作员复位......") Then Exit Function
    MsgBox "操作员复位成功！", vbInformation, gstrSysName
End Function

Private Function 昭通_手工冲销门诊单据() As Boolean
    Dim arrData
    Dim strData As String
    
    strData = InputBox("请输入就诊编号与个人帐户，中间以;号分隔：", "手工冲销门诊单据", "1010111;50.00")
    If Trim(strData) = "" Then Exit Function
    If InStr(1, strData, ";") = 0 Then
        MsgBox "格式错误，正确格式：就诊编号;个人帐户，如：1010111;50.00", vbInformation, gstrSysName
        Exit Function
    End If
    If Not IsNumeric(Split(strData, ";")(1)) Then
        MsgBox "格式错误，正确格式：就诊编号;个人帐户，如：1010111;50.00", vbInformation, gstrSysName
        Exit Function
    End If
    arrData = Split(strData, ";")
    
    If Not frmConn昭通.Execute("I220", 0, arrData(0) & vbTab & arrData(1), "手工冲销门诊单据......") Then Exit Function
    MsgBox "手工冲销门诊单据成功！", vbInformation, gstrSysName
End Function

Private Function 昭通_申报药品() As Boolean
    '在HIS中提取标准目录中的数据进行项目对照，然后统一申报，成功后更新执行库
    Dim rsTemp As New ADODB.Recordset, rs项目 As New ADODB.Recordset, strTemp As String
    
    Set rs项目 = gcnOracle.Execute("Select a.收费细目id,a.项目编码,b.计算单位,b.名称,b.类别 From 保险支付项目 a,收费细目 b Where 险类=103 and nvl(附注,'未审批')='未审批' And a.收费细目id=b.id And (b.类别='5' or b.类别='6' or b.类别='7')")
    If rs项目.EOF Then
        MsgBox "没有项目需要申报", vbInformation, "申报"
        Exit Function
    End If
    
    While Not rs项目.EOF
        Set rsTemp = gcnOracle.Execute("Select * from 药品目录 Where 药品ID=" & rs项目!收费细目ID)
        strTemp = LeftStr(rs项目!名称, 28) & vbTab & LeftStr(Nvl(rs项目!计算单位, " "), 4) & _
            vbTab & LeftStr(Nvl(rsTemp!规格, " "), 16) & vbTab & _
            LeftStr(Nvl(rsTemp!药品来源, " "), 12) & vbTab & rsTemp!指导零售价 & vbTab & _
            LeftStr(Nvl(rsTemp!产地, " "), 30) & vbTab & LeftStr(Nvl(rsTemp!批准文号, " "), 30)
        Set rsTemp = gcn昭通.Execute("Select * From tab_syml Where dm='" & rs项目!项目编码 & "'")
        strTemp = rs项目!收费细目ID & vbTab & " " & vbTab & rsTemp!dm & vbTab & _
            rsTemp!lb & vbTab & strTemp
        
        frmConn昭通.Execute "I110", 1, strTemp, "正在进行药品申报......"
        rs项目.MoveNext
    Wend
    昭通_申报药品 = True
End Function

Private Function 昭通_下载药品执行库() As Boolean
    '下载药品目录执行库
    Dim str更新状态 As String, rsTemp As New ADODB.Recordset, strData() As String, lngLoop As Long
    Dim strTemp As String
    On Error GoTo errHandle

    '以前的方式不可取，改为使用11功能号，先将所有项目更新为未审核，再根据返回数据更新，只要有就肯定是已审批
    gcnOracle.BeginTrans
    If Not frmConn昭通.Execute("I110", 11, "", "正在获取药品目录执行库数据......") Then gcnOracle.RollbackTrans: Exit Function

    Call ShowWindow(frmConn昭通.hwnd, 9)
    DoEvents

    gstrSQL = "Update 保险支付项目 " & _
             " Set 附注='未审批' " & _
             " Where 险类=103 And 是否医保=1 And 项目编码 is not null " & _
             " And 收费细目ID In ( " & _
             "     Select Id From 收费细目 Where 类别 In ('5','6','7'))"
    gcnOracle.Execute gstrSQL
    For lngLoop = 1 To frmConn昭通.mlngRows
        If frmConn昭通.Query(lngLoop - 1, 1, "正在更新数据(" & lngLoop & "/" & frmConn昭通.mlngRows & ")......") = False Then gcnOracle.RollbackTrans: Exit Function
        strTemp = Replace(frmConn昭通.strReturnInfo, "'", "''")
        gcnOracle.Execute "update 保险支付项目 set 附注='已审批' Where 收费细目ID=" & Split(strTemp, vbTab)(0)
    Next
    Call ShowWindow(frmConn昭通.hwnd, 0)
    
    gcnOracle.CommitTrans
    昭通_下载药品执行库 = True
    Exit Function

errHandle:
    If MsgBox("更新项目时发生错误：" & vbCrLf & Err.Description & vbCrLf & "是否重试？", vbInformation + vbRetryCancel, "错误") = vbRetry Then
        Err.Clear
        Resume
    End If
    Call ShowWindow(frmConn昭通.hwnd, 0)
    gcnOracle.RollbackTrans
End Function

Private Function 昭通_下载标准药品目录() As Boolean
    '下载标准药品目录
    Dim str更新状态 As String, rsTemp As New ADODB.Recordset, strData() As String, lngLoop As Long
    Dim strTemp As String
    Dim arrData
    
    If Not frmConn昭通.Execute("I100", 0, "", "获取医保数据更新状态......") Then Exit Function
    If frmConn昭通.Query(0, 0) = False Then Exit Function
    If frmConn昭通.strReturnInfo = "" Then
        If MsgBox("不能取得医保数据更新状态，是否继续？", vbQuestion + vbYesNo, "更新") = vbNo Then
            Exit Function
        Else
            str更新状态 = ""
        End If
    Else
        str更新状态 = Split(frmConn昭通.strReturnInfo, vbTab)(0)
    End If
    
    Set rsTemp = gcn昭通.Execute("Select * From TAB_UPDATE")
    If rsTemp.EOF Then
        gcn昭通.Execute "Insert Into TAB_UPDATE Values (NULL,'" & str更新状态 & "',NULL)"
    Else
        If IsNull(rsTemp!BZML) Then
            gcn昭通.Execute "Update TAB_UPDATE Set BZML='" & str更新状态 & "'"
        ElseIf rsTemp!BZML = str更新状态 Then
            If MsgBox("从上次下载以来，标准药品目录未进行更新，是否重新下载？", vbYesNo + vbQuestion, "下载标准药品目录") = vbNo Then
                Exit Function
            End If
        Else
            gcn昭通.Execute "Update TAB_UPDATE Set BZML='" & str更新状态 & "'"
        End If
    End If
    
    If Not frmConn昭通.Execute("I100", 3, "", "正在获取标准药品目录数据......") Then Exit Function
    
    On Error GoTo errHandle
    gcn昭通.BeginTrans
    gcn昭通.Execute "Delete From tab_syml"
    Call ShowWindow(frmConn昭通.hwnd, 9)
    DoEvents
    For lngLoop = 1 To frmConn昭通.mlngRows
        If frmConn昭通.Query(lngLoop - 1, 1, "正在更新数据(" & lngLoop & "/" & frmConn昭通.mlngRows & ")......") = False Then
            gcn昭通.RollbackTrans
            Exit Function
        End If
        strTemp = frmConn昭通.strReturnInfo
        arrData = Split(strTemp, vbTab)
        strTemp = Replace(arrData(11), "'", "") 'rq
        If Trim(strTemp) <> "" Then
            strTemp = "to_date('" & strTemp & "','yyyyMMdd')"
        Else
            strTemp = "''"
        End If
        gcn昭通.Execute "Insert Into tab_syml (dl,ty,dm,tm,sm,lb,py,dw,dj,jx,gg,rq,zt,xd,xj,cs) " & _
            " values (" & _
            "'" & Replace(arrData(0), "'", "''") & "','" & Replace(arrData(1), "'", "''") & "'," & _
            "'" & Replace(arrData(2), "'", "''") & "','" & Replace(arrData(3), "'", "''") & "'," & _
            "'" & Replace(arrData(4), "'", "''") & "','" & Replace(arrData(5), "'", "''") & "'," & _
            "'" & Replace(arrData(6), "'", "''") & "','" & Replace(arrData(7), "'", "''") & "'," & _
            "'" & Replace(arrData(8), "'", "''") & "','" & Replace(arrData(9), "'", "''") & "'," & _
            "'" & Replace(arrData(10), "'", "''") & "'," & strTemp & "," & _
            "'" & Replace(arrData(12), "'", "''") & "','" & Replace(arrData(13), "'", "''") & "'," & _
            "'" & Replace(arrData(14), "'", "''") & "','" & Replace(arrData(15), "'", "''") & "')"
    Next
    Call ShowWindow(frmConn昭通.hwnd, 0)
    gcn昭通.CommitTrans
    昭通_下载标准药品目录 = True
    Exit Function
    
errHandle:
    If MsgBox("更新项目时发生错误：" & vbCrLf & Err.Description & vbCrLf & "是否重试？", vbInformation + vbRetryCancel, "错误") = vbRetry Then
        Err.Clear
        Resume
    End If
    Call ShowWindow(frmConn昭通.hwnd, 0)
    gcn昭通.RollbackTrans
End Function

Private Function 昭通_下载材料项目() As Boolean
    '下载材料项目
    Dim str更新状态 As String, rsTemp As New ADODB.Recordset, strData() As String, lngLoop As Long
    Dim strTemp As String
    
    If Not frmConn昭通.Execute("I100", 0, "", "获取医保数据更新状态......") Then Exit Function
    If frmConn昭通.Query(0, 0) = False Then Exit Function
    
    If frmConn昭通.strReturnInfo = "" Then
        If MsgBox("不能取得医保数据更新状态，是否继续？", vbQuestion + vbYesNo, "更新") = vbNo Then
            Exit Function
        Else
            str更新状态 = ""
        End If
    Else
        str更新状态 = Split(frmConn昭通.strReturnInfo, vbTab)(0)
    End If
    
    Set rsTemp = gcn昭通.Execute("Select * From TAB_UPDATE")
    If rsTemp.EOF Then
        gcn昭通.Execute "Insert Into TAB_UPDATE Values (NULL,'" & str更新状态 & "',NULL)"
    Else
        If IsNull(rsTemp!CLXM) Then
            gcn昭通.Execute "Update TAB_UPDATE Set CLXM='" & str更新状态 & "'"
        ElseIf rsTemp!CLXM = str更新状态 Then
            If MsgBox("从上次下载以来，材料项目未进行更新，是否重新下载？", vbYesNo + vbQuestion, "下载材料项目") = vbNo Then
                Exit Function
            End If
        Else
            gcn昭通.Execute "Update TAB_UPDATE Set CLXM='" & str更新状态 & "'"
        End If
    End If
    
    If Not frmConn昭通.Execute("I100", 2, "", "正在获取材料项目数据......") Then Exit Function
'    If frmConn昭通.Query(0, 0) = False Then Exit Sub
'
'    strData = Split(frmConn昭通.strReturnInfo, Chr(10))
    
    On Error GoTo errHandle
    gcn昭通.BeginTrans
    gcn昭通.Execute "Delete From tab_fwcl Where lb In (31,32,33,51,52,53)"
    Call ShowWindow(frmConn昭通.hwnd, 9)
    DoEvents
    For lngLoop = 1 To frmConn昭通.mlngRows
        If frmConn昭通.Query(lngLoop - 1, 1, "正在更新数据(" & lngLoop & "/" & frmConn昭通.mlngRows & ")......") = False Then
            gcn昭通.RollbackTrans
            Exit Function
        End If
        strTemp = frmConn昭通.strReturnInfo
        gcn昭通.Execute "Insert Into tab_fwcl (dm,kc,lb,mc,dw,dj,cx) values ('" & _
            Split(strTemp, vbTab)(0) & "','" & _
            Split(strTemp, vbTab)(1) & "'," & _
            Split(strTemp, vbTab)(2) & ",'" & _
            Split(strTemp, vbTab)(3) & "','" & _
            Split(strTemp, vbTab)(4) & "'," & _
            Split(strTemp, vbTab)(5) & "," & _
            IIf(Split(strTemp, vbTab)(6) = " ", "NULL", Split(strTemp, vbTab)(6)) & ")"
    Next
    Call ShowWindow(frmConn昭通.hwnd, 0)

    gcn昭通.CommitTrans
    昭通_下载材料项目 = True
    Exit Function
    
errHandle:
    If MsgBox("更新项目时发生错误：" & vbCrLf & Err.Description & vbCrLf & "是否重试？", vbInformation + vbRetryCancel, "错误") = vbRetry Then
        Err.Clear
        Resume
    End If
    Call ShowWindow(frmConn昭通.hwnd, 0)
    gcn昭通.RollbackTrans
End Function

Private Function 昭通_下载服务项目() As Boolean
    '下载服务项目
    Dim str更新状态 As String, rsTemp As New ADODB.Recordset, strData() As String, lngLoop As Long
    Dim strTemp As String
    If Not frmConn昭通.Execute("I100", 0, "", "获取医保数据更新状态......") Then Exit Function
    If frmConn昭通.Query(0, 0) = False Then Exit Function
    If frmConn昭通.strReturnInfo = "" Then
        If MsgBox("不能取得医保数据更新状态，是否继续？", vbQuestion + vbYesNo, "更新") = vbNo Then
            Exit Function
        Else
            str更新状态 = ""
        End If
    Else
        str更新状态 = Split(frmConn昭通.strReturnInfo, vbTab)(0)
    End If
    
    Set rsTemp = gcn昭通.Execute("Select * From TAB_UPDATE")
    If rsTemp.EOF Then
        gcn昭通.Execute "Insert Into TAB_UPDATE Values ('" & str更新状态 & "',NULL,NULL)"
    Else
        If IsNull(rsTemp!FWXM) Then
            gcn昭通.Execute "Update TAB_UPDATE Set FWXM='" & str更新状态 & "'"
        ElseIf rsTemp!FWXM = str更新状态 Then
            If MsgBox("从上次下载以来，服务项目未进行更新，是否重新下载？", vbYesNo + vbQuestion, "下载服务项目") = vbNo Then
                Exit Function
            End If
        Else
            gcn昭通.Execute "Update TAB_UPDATE Set FWXM='" & str更新状态 & "'"
        End If
    End If
    
    If Not frmConn昭通.Execute("I100", 1, "", "正在获取服务项目数据......") Then Exit Function
'    If frmConn昭通.Query(0, 0) = False Then Exit Sub
'    strData = Split(frmConn昭通.strReturnInfo, Chr(10))
    
    On Error GoTo errHandle
    gcn昭通.BeginTrans
    gcn昭通.Execute "Delete From tab_fwcl Where lb In (20,21,22,23,24,25,40)"
    Call ShowWindow(frmConn昭通.hwnd, 9)
    DoEvents
    For lngLoop = 1 To frmConn昭通.mlngRows
        If frmConn昭通.Query(lngLoop - 1, 1, "正在更新数据(" & lngLoop & "/" & frmConn昭通.mlngRows & ")......") = False Then
            gcn昭通.RollbackTrans
            Exit Function
        End If
        strTemp = frmConn昭通.strReturnInfo
        gcn昭通.Execute "Insert Into tab_fwcl (dm,kc,lb,mc,dw,dj,cx) values ('" & _
            Split(strTemp, vbTab)(0) & "','" & _
            Split(strTemp, vbTab)(1) & "'," & _
            Split(strTemp, vbTab)(2) & ",'" & _
            Replace(Split(strTemp, vbTab)(3), "'", "") & "','" & _
            Split(strTemp, vbTab)(4) & "'," & _
            Split(strTemp, vbTab)(5) & "," & _
            IIf(Split(strTemp, vbTab)(6) = " ", "NULL", Split(strTemp, vbTab)(6)) & ")"
    Next
    Call ShowWindow(frmConn昭通.hwnd, 0)
    gcn昭通.CommitTrans
    昭通_下载服务项目 = True
    Exit Function
    
errHandle:
    If MsgBox("更新项目时发生错误：" & vbCrLf & Err.Description & vbCrLf & "是否重试？", vbInformation + vbRetryCancel, "错误") = vbRetry Then
        Err.Clear
        Resume
    End If
    Call ShowWindow(frmConn昭通.hwnd, 0)
    gcn昭通.RollbackTrans
End Function

Private Function 涪陵工具包(ByVal lng功能号 As Long) As Boolean
    Dim strJZBH As String
    Const C涪陵_手工冲销门诊结算 As Integer = 200
    
    If Not mclsInsure.InitInsure(gcnOracle, TYPE_涪陵) Then Exit Function
    Select Case lng功能号
    Case C涪陵_手工冲销门诊结算
        '处理意外中止导致的中心有数据而HIS无数据的情况，由操作员到前置机查其结算编号，在此处录入即可完成门诊结算作废
        涪陵工具包 = 涪陵_手工冲销门诊结算
    End Select
End Function

Private Function 涪陵_手工冲销门诊结算() As Boolean
    Dim str就诊编号 As String
    Dim blnReturn As Boolean
    
    If gstr医保机构编码 = "" Then
        MsgBox "准备读取医保机构编码，请插入系统卡或病人卡", vbInformation, gstrSysName
CheckCard:
        initType
        blnReturn = fl_getybjgbm(gstrOutPara)
        TrimType
        If blnReturn = False Then
            If MsgBox(gstrOutPara.errtext, vbRetryCancel, gstrSysName) = vbRetry Then
                GoTo CheckCard
            Else
                Exit Function
            End If
        End If
        gstr医保机构编码 = gstrOutPara.out1
        gstr医院编码 = gstrOutPara.out2
    End If
        
    On Error GoTo errHandle
    str就诊编号 = InputBox("请录入就诊编号：", "根据就诊编号作废门诊结算")
    If Trim(str就诊编号) = "" Then
        MsgBox "请诊编号为空，无法完成门诊结算作废！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '调用接口数冲销
    initType
    blnReturn = fl_canrollback(gstr医保机构编码, gstr医院编码, str就诊编号, gstrOutPara)
    TrimType
    If blnReturn = False Then
        MsgBox "判断是否可以冲销时，医保端返回以下信息，退费不能继续。" & Chr(13) & Chr(10) & gstrOutPara.errtext, vbInformation, gstrSysName
        Exit Function
    End If
    initType
    blnReturn = fl_rollbackcalc(gstr医保机构编码, gstr医院编码, str就诊编号, "0", gstrOutPara)
    TrimType
    If blnReturn = False Then
        MsgBox gstrOutPara.errtext, vbInformation, gstrSysName
        Exit Function
    End If

    涪陵_手工冲销门诊结算 = True
    Exit Function
errHandle:
    MsgBox "错误发生在[手工冲销门诊结算]工具中，错误信息：" & Chr(13) & Chr(10) & Err.Description
End Function



