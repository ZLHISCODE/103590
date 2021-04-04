VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmFinanceSupervisePersonList 
   BorderStyle     =   0  'None
   Caption         =   "人员信息列表"
   ClientHeight    =   8865
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11655
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   11655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picBalance 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   105
      ScaleHeight     =   570
      ScaleWidth      =   11625
      TabIndex        =   12
      Top             =   6915
      Width           =   11655
      Begin VB.Label lblBalance 
         Caption         =   "当前暂存金:"
         Height          =   945
         Left            =   30
         TabIndex        =   13
         Top             =   75
         Width           =   7935
      End
   End
   Begin VB.PictureBox picPersonPage 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2520
      Left            =   285
      ScaleHeight     =   2520
      ScaleWidth      =   3435
      TabIndex        =   9
      Top             =   3810
      Width           =   3435
      Begin XtremeSuiteControls.TabControl tbPage 
         Height          =   1605
         Left            =   -15
         TabIndex        =   10
         Top             =   -30
         Width           =   2865
         _Version        =   589884
         _ExtentX        =   5054
         _ExtentY        =   2831
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picOtherList 
      BorderStyle     =   0  'None
      Height          =   1140
      Left            =   195
      ScaleHeight     =   1140
      ScaleWidth      =   2745
      TabIndex        =   4
      Top             =   2055
      Width           =   2745
      Begin VB.TextBox txtOtherPerson 
         ForeColor       =   &H80000000&
         Height          =   315
         Left            =   675
         TabIndex        =   8
         Tag             =   "输入简码或汉字或编号"
         Text            =   "输入简码或汉字或编号"
         Top             =   0
         Width           =   2310
      End
      Begin MSComctlLib.ListView lvwOther_S 
         Height          =   825
         Left            =   165
         TabIndex        =   11
         Top             =   555
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   1455
         View            =   2
         Arrange         =   1
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         OLEDropMode     =   1
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ilsbig"
         SmallIcons      =   "ilssmall"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OLEDropMode     =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "姓名"
            Text            =   "姓名"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "编号"
            Text            =   "编号"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Key             =   "简码"
            Text            =   "简码"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Key             =   "所属部门"
            Text            =   "所属部门"
            Object.Width           =   4304
         EndProperty
      End
      Begin VB.Label lblOtherPerson 
         AutoSize        =   -1  'True
         Caption         =   "收费员"
         Height          =   210
         Left            =   0
         TabIndex        =   7
         Top             =   30
         Width           =   630
      End
   End
   Begin VB.PictureBox picGroupList 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1140
      Left            =   135
      ScaleHeight     =   1140
      ScaleWidth      =   2745
      TabIndex        =   2
      Top             =   1410
      Width           =   2745
      Begin MSComctlLib.ListView lvwGroup_S 
         Height          =   825
         Left            =   -15
         TabIndex        =   3
         Top             =   0
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   1455
         Arrange         =   2
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         OLEDropMode     =   1
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ilsbig"
         SmallIcons      =   "ilssmall"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OLEDropMode     =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "组名称"
            Object.Tag             =   "组名称"
            Text            =   "组名称"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "编码"
            Object.Tag             =   "编码"
            Text            =   "编码"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Key             =   "简码"
            Object.Tag             =   "简码"
            Text            =   "简码"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Key             =   "负责人姓名"
            Object.Tag             =   "负责人姓名"
            Text            =   "负责人姓名"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Key             =   "说明"
            Object.Tag             =   "说明"
            Text            =   "说明"
            Object.Width           =   4304
         EndProperty
      End
   End
   Begin VB.PictureBox picPersonList 
      BorderStyle     =   0  'None
      Height          =   1755
      Left            =   390
      ScaleHeight     =   1755
      ScaleWidth      =   2700
      TabIndex        =   0
      Top             =   75
      Width           =   2700
      Begin VB.TextBox txtChargePerson 
         ForeColor       =   &H80000000&
         Height          =   315
         Left            =   675
         TabIndex        =   6
         Tag             =   "输入简码或汉字或编号"
         Text            =   "输入简码或汉字或编号"
         Top             =   45
         Width           =   2310
      End
      Begin MSComctlLib.ListView lvwPerson_S 
         Height          =   825
         Left            =   90
         TabIndex        =   1
         Top             =   495
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   1455
         View            =   2
         Arrange         =   1
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         OLEDropMode     =   1
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ilsbig"
         SmallIcons      =   "ilssmall"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OLEDropMode     =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "姓名"
            Text            =   "姓名"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "编号"
            Text            =   "编号"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Key             =   "简码"
            Text            =   "简码"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Key             =   "所属部门"
            Text            =   "所属部门"
            Object.Width           =   4304
         EndProperty
      End
      Begin VB.Label lblChargePerson 
         AutoSize        =   -1  'True
         Caption         =   "收费员"
         Height          =   210
         Left            =   0
         TabIndex        =   5
         Top             =   75
         Width           =   630
      End
   End
   Begin MSComctlLib.ImageList ilssmall 
      Left            =   4170
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFinanceSupervisePersonList.frx":0000
            Key             =   "Man"
            Object.Tag             =   "Man"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFinanceSupervisePersonList.frx":059A
            Key             =   "Woman"
            Object.Tag             =   "Woman"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFinanceSupervisePersonList.frx":0B34
            Key             =   "Group"
            Object.Tag             =   "Group"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsbig 
      Left            =   4365
      Top             =   195
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFinanceSupervisePersonList.frx":10CE
            Key             =   "Man"
            Object.Tag             =   "Man"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFinanceSupervisePersonList.frx":19A8
            Key             =   "Woman"
            Object.Tag             =   "Woman"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFinanceSupervisePersonList.frx":2282
            Key             =   "Group"
            Object.Tag             =   "Group"
         EndProperty
      EndProperty
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmFinanceSupervisePersonList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum mPgIndex
    EM_PG_收费员 = 250101
    EM_PG_财务组 = 250102
    EM_PG_其他人员 = 250103
End Enum
Private Enum mPaneIndex
    EM_PN_收费员列表 = 1
    EM_PN_明细列表 = 2
    EM_PN_暂存列表 = 3
End Enum
Private WithEvents mfrmList As frmFinaceSuperviseCollectList
Attribute mfrmList.VB_VarHelpID = -1
Private mfrmPersonOther As frmFinanceSupervisePersonOthers

Private mlngModule As Long, mstrPrivs As String
Private mrsChargePerson As ADODB.Recordset '收费员记录集
Private mrsOtherPerson As ADODB.Recordset   '其他人员记录集
Private mrsGroup As ADODB.Recordset
Private mstrSelPerson As String '上次选择的收费人员
Private mstrSelGroup As String '上次选择的组人员
Private mstrSelOther As String '上次选择的其他人员
Private mcbsMain As Object
Private mblnNotBrush As Boolean '不刷新数据
Private mobjDetailPane As Pane
Private Function LoadPersonFromLvw(ByVal objLvw As ListView, ByVal rsTemp As ADODB.Recordset) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载收费员给控件
    '入参:objLvw-加载的数据
    '       rsTemp-收费员集(ID,编号,姓名,简码,所属部门)
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-09-24 11:40:05
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objItem As ListItem, strIcon As String
    On Error GoTo errHandle
    '全部加载
    With objLvw.ListItems
        .Clear
        If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
        Do While Not rsTemp.EOF
            If Nvl(rsTemp!性别) Like "*男*" Then
                strIcon = "Man"
            Else
                strIcon = "Woman"
            End If
            Set objItem = .Add(, "K" & Nvl(rsTemp!编号), Nvl(rsTemp!姓名), strIcon, strIcon)
            objItem.SubItems(1) = Nvl(rsTemp!编号)
            objItem.SubItems(2) = Nvl(rsTemp!简码)
            objItem.SubItems(3) = Nvl(rsTemp!所属部门)
            objItem.Tag = Nvl(rsTemp!ID)
            rsTemp.MoveNext
        Loop
    End With
    LoadPersonFromLvw = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
End Function

Private Function LoadPerson(Optional blnFilter As Boolean = False) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载收费员信息
    '入参:blnFilter-是否进行过滤
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-09-23 11:59:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsReturn As ADODB.Recordset, rsTemp As ADODB.Recordset
    Dim intInputType As Integer '0-输入的是全数字,1-输入的是全字母,2-其他
    Dim strCompents As String '匹配串
    Dim i As Long, intIdx As Integer, iCount As Integer
    Dim strText As String, strResult As String, strFilter As String
    Dim strSQL As String, strIcon As String
    On Error GoTo errHandle
    
    If zlStr.IsHavePrivs(mstrPrivs, "收费员收款") = False Then LoadPerson = True: Exit Function
    
    '读取收费员信息
    If blnFilter = False Or mrsChargePerson Is Nothing Then
        strSQL = "" & _
        "   Select distinct A.ID,A.编号,A.姓名,A.简码,M.名称 as 所属部门,a.性别" & _
        "   From 人员表 A,人员性质说明 B, 部门人员 C,部门表 M" & _
        "   Where A.id = B.人员ID And B.人员性质 In ('门诊挂号员','门诊收费员','预交收款员','住院结帐员','入院登记员','发卡登记人')  " & _
        "               And A.ID=C.人员ID and C.部门ID=M.ID(+) And C.缺省(+)=1 " & _
        "               And (A.撤档时间 Is Null Or A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
        "               And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & vbNewLine & _
        "   Order By 编号"
        Set mrsChargePerson = zlDatabase.OpenSQLRecord(strSQL, "获取收费员信息")
    End If
    mrsChargePerson.Filter = 0
    strText = UCase(txtChargePerson.Text)
    
    If txtChargePerson.Text = txtChargePerson.Tag Or strText = "" Then
        '全部加载
        LoadPerson = LoadPersonFromLvw(lvwPerson_S, mrsChargePerson)
        Exit Function
    End If
    
    strCompents = gstrLike & strText & "%"
    If IsNumeric(strText) Then '1.输入的是全数字
    ElseIf zlCommFun.IsCharAlpha(strText) Then '1-输入的是全字母
        mrsChargePerson.Filter = "简码 like '" & gstrLike & strText & "%'"
        LoadPerson = LoadPersonFromLvw(lvwPerson_S, mrsChargePerson)
        Exit Function
    Else
        intInputType = 2   '2-其他
        mrsChargePerson.Filter = "姓名 like '" & strText & "%'"
        LoadPerson = LoadPersonFromLvw(lvwPerson_S, mrsChargePerson)
        Exit Function
    End If
    
    '输入的是全数字
    Set rsTemp = zlDatabase.zlCopyDataStructure(mrsChargePerson)
    With mrsChargePerson
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            '如果输入的数字,需要检查:
            '1.编号输入值相等,主要输入如:12 匹配000012这种情况
            '2.输入的数字,则认为是编码,只能左匹配,比如输入12匹配00001201或120001等
            '主要是检查输入的内容与编号完全相同,则直接就定位到该姓名
            If Nvl(!编号) = strText Then
                Call zlDatabase.zlInsertCurrRowData(mrsChargePerson, rsTemp): Exit Do
            End If
            
            '1.编号输入值相等,主要输入如:12 匹配000012这种情况,因为这种情况有很多:如0012,012,000012等
            If Val(Nvl(!编号)) = Val(strText) Then
                Call zlDatabase.zlInsertCurrRowData(mrsChargePerson, rsTemp)
            End If
            '2.输入的数字,则认为是编码,只能左匹配,比如输入12匹配00001201或120001等
             If Val(Nvl(!编号)) Like strText & "*" Then Call zlDatabase.zlInsertCurrRowData(mrsChargePerson, rsTemp)
            mrsChargePerson.MoveNext
        Loop
    End With
    LoadPerson = LoadPersonFromLvw(lvwPerson_S, rsTemp)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function LoadOtherPerson(Optional blnFilter As Boolean = False) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载其他收费员信息
    '入参:blnFilter-是否进行过滤
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-09-23 11:59:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsReturn As ADODB.Recordset, rsTemp As ADODB.Recordset
    Dim intInputType As Integer '0-输入的是全数字,1-输入的是全字母,2-其他
    Dim strCompents As String '匹配串
    Dim strText As String, strResult As String, strFilter As String
    Dim strSQL As String, strIcon As String
    
    On Error GoTo errHandle
    If zlStr.IsHavePrivs(mstrPrivs, "其他人员收款") = False Then LoadOtherPerson = True: Exit Function
    '读取收费员信息
    If blnFilter = False Or mrsOtherPerson Is Nothing Then

        strSQL = " " & _
        "   Select Distinct a.Id, a.编号, a.姓名, a.简码, m.名称 As 所属部门, a.性别 " & _
        "   From 人员缴款余额 A1, 人员表 A, 部门人员 C, 部门表 M " & _
        "   Where A1.收款员 = a.姓名 And A1.性质 = 1    " & _
        "         And not exists(select 1 From  人员性质说明 B where a.ID=b.人员ID And  b.人员性质  In  ('门诊挂号员', '门诊收费员', '预交收款员', '住院结帐员', '入院登记员', '发卡登记人')) " & _
        "         And a.Id = c.人员id And c.部门id = m.Id(+) And  c.缺省(+) = 1" & _
        "         And (A.撤档时间 Is Null Or A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
        "         And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & vbNewLine & _
        "   Order By 编号"
        Set mrsOtherPerson = zlDatabase.OpenSQLRecord(strSQL, "获取其他收费员信息")
    End If
    mrsOtherPerson.Filter = 0
    strText = UCase(txtOtherPerson.Text)
    If txtOtherPerson.Text = txtOtherPerson.Tag Or strText = "" Then
        '全部加载
        LoadOtherPerson = LoadPersonFromLvw(lvwOther_S, mrsOtherPerson)
        Exit Function
    End If
    strCompents = gstrLike & strText & "%"
    If IsNumeric(strText) Then '1.输入的是全数字
    ElseIf zlCommFun.IsCharAlpha(strText) Then '1-输入的是全字母
        mrsOtherPerson.Filter = "简码 like '" & gstrLike & strText & "%'"
        LoadOtherPerson = LoadPersonFromLvw(lvwOther_S, mrsOtherPerson)
        Exit Function
    Else
        intInputType = 2   '2-其他
        mrsOtherPerson.Filter = "姓名 like '" & strText & "%'"
        LoadOtherPerson = LoadPersonFromLvw(lvwOther_S, mrsOtherPerson)
        Exit Function
    End If
    
    '输入的是全数字
    Set rsTemp = zlDatabase.zlCopyDataStructure(mrsOtherPerson)
    With mrsOtherPerson
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            '如果输入的数字,需要检查:
            '1.编号输入值相等,主要输入如:12 匹配000012这种情况
            '2.输入的数字,则认为是编码,只能左匹配,比如输入12匹配00001201或120001等
            '主要是检查输入的内容与编号完全相同,则直接就定位到该姓名
            If Nvl(!编号) = strText Then
                Call zlDatabase.zlInsertCurrRowData(mrsOtherPerson, rsTemp): Exit Do
            End If
            '1.编号输入值相等,主要输入如:12 匹配000012这种情况,因为这种情况有很多:如0012,012,000012等
            If Val(Nvl(!编号)) = Val(strText) Then
                Call zlDatabase.zlInsertCurrRowData(mrsOtherPerson, rsTemp)
            End If
            '2.输入的数字,则认为是编码,只能左匹配,比如输入12匹配00001201或120001等
             If Val(Nvl(!编号)) Like strText & "*" Then Call zlDatabase.zlInsertCurrRowData(mrsOtherPerson, rsTemp)
            mrsOtherPerson.MoveNext
        Loop
    End With
    LoadOtherPerson = LoadPersonFromLvw(lvwOther_S, rsTemp)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function LoadGroup() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载财务组数据
    '入参:  blnFilter-是否过滤
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-09-24 12:16:41
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
   Dim objItem As ListItem, strSQL As String
   Dim rsGroup As ADODB.Recordset
   Dim str负责人 As String
    On Error GoTo errHandle
    
    If zlStr.IsHavePrivs(mstrPrivs, "财务组收款") = False Then LoadGroup = True: Exit Function
    '读取财务组
    strSQL = " " & _
    "   Select a.Id As 编码, a.组名称, a.简码, b.姓名 As 组负责人,A.负责人id ,A.说明" & _
    "   From 财务缴款分组 A, 人员表 B " & _
    "   Where a.负责人id = b.Id And Nvl(a.删除日期, To_Date('3000-01-01', 'yyyy-mm-dd')) >= To_Date('3000-01-01', 'yyyy-mm-dd') And " & _
    "         (b.撤档时间 Is Null Or b.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) And (b.站点 = 'A' Or b.站点 Is Null) " & _
    "   Order By a.组名称"
    Set mrsGroup = zlDatabase.OpenSQLRecord(strSQL, "获取财务组信息")
    '全部加载
    With lvwGroup_S.ListItems
        .Clear
        If mrsGroup.RecordCount <> 0 Then mrsGroup.MoveFirst
        Do While Not mrsGroup.EOF
            Set objItem = .Add(, "K" & Nvl(mrsGroup!编码), Nvl(mrsGroup!组名称), "Group", "Group")
            objItem.SubItems(1) = Nvl(mrsGroup!编码)
            objItem.SubItems(2) = Nvl(mrsGroup!简码)
            strSQL = "Select B.姓名 From 财务缴款分组 A,人员表 B Where (A.删除日期 Is Null or A.删除日期 Between Sysdate And to_date('3000-01-01','YYYY-MM-DD')) And A.负责人ID = B.ID And A.ID = [1]"
            strSQL = strSQL & " Union Select C.姓名 From 财务组组长构成 A,财务缴款分组 B,人员表 C Where A.组ID=B.ID And A.组长ID=C.ID And B.ID = [1] And (B.删除日期 Is Null or B.删除日期 Between Sysdate And to_date('3000-01-01','YYYY-MM-DD'))"
            Set rsGroup = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Nvl(mrsGroup!编码)))
            str负责人 = ""
            Do While Not rsGroup.EOF
                str负责人 = str负责人 & "," & rsGroup!姓名
                rsGroup.MoveNext
            Loop
            If str负责人 <> "" Then str负责人 = Mid(str负责人, 2)
            objItem.SubItems(3) = str负责人
            objItem.SubItems(4) = Nvl(mrsGroup!说明)
            objItem.Tag = Val(Nvl(mrsGroup!负责人id))
            mrsGroup.MoveNext
        Loop
    End With
    LoadGroup = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Function
Public Sub zlInitVar(ByVal lngModule As Long, ByVal strPrivs As String, ByRef cbsMain As Object)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化相关变量
    '入参:lngModule-模块号
    '       strPrivs-权限串
    '编制:刘兴洪
    '日期:2013-09-09 14:41:46
    '说明:加载窗体后,立即调用
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlngModule = lngModule: mstrPrivs = strPrivs
    Set mcbsMain = cbsMain
End Sub

Private Sub Form_Load()
    mstrPrivs = gstrPrivs: mlngModule = glngModul
    Call InitPage: Call InitPanel
    Call LoadGroup: Call LoadPerson(False): Call LoadOtherPerson(False)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If Not mfrmList Is Nothing Then Unload mfrmList
    Set mfrmList = Nothing
    If Not mfrmPersonOther Is Nothing Then Unload mfrmPersonOther
    Set mfrmPersonOther = Nothing
End Sub

Private Sub lvwGroup_S_GotFocus()
    mstrSelGroup = ""
End Sub
Private Sub lvwGroup_S_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If Item.Text = mstrSelGroup Then Exit Sub
    mstrSelGroup = Item.Text
    Call LoadLocalePersonDetailData
End Sub

Private Sub lvwGroup_S_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then Exit Sub
    Call ShowPopup
End Sub

Private Sub lvwOther_S_GotFocus()
    mstrSelOther = ""
End Sub

Private Sub lvwOther_S_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If Item.Text = mstrSelOther Then Exit Sub
    mstrSelOther = Item.Text
    Call LoadBalance(mstrSelOther)
    Call LoadLocalePersonDetailData
End Sub
Private Sub lvwOther_S_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then Exit Sub
    Call ShowPopup
End Sub
Private Sub lvwPerson_S_GotFocus()
    mstrSelPerson = ""
End Sub
Private Sub lvwPerson_S_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If Item.Text = mstrSelPerson Then Exit Sub
    mstrSelPerson = Item.Text
    Call LoadLocalePersonDetailData
End Sub
Private Sub lvwPerson_S_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then Exit Sub
    Call ShowPopup
End Sub
Private Sub mfrmList_PersonChange(ByVal strPerson As String, Cancel As Boolean)
    '人员改变时,需要定位到指定的人员上
    Dim objItem As ListItem
    If Val(tbPage.Selected.Tag) = EM_PG_其他人员 Then Cancel = True: Exit Sub
    mblnNotBrush = True
    If Val(tbPage.Selected.Tag) = EM_PG_财务组 Then
        For Each objItem In lvwGroup_S.ListItems
            If InStr("," & objItem.SubItems(3) & ",", "," & strPerson & ",") > 0 Then
                Call LoadBalance(strPerson): mblnNotBrush = False
                objItem.Selected = True: Exit Sub
            End If
        Next
        mblnNotBrush = False
        Cancel = True: Exit Sub
    End If
    '收费员轧帐
    For Each objItem In lvwPerson_S.ListItems
        If objItem.Text = strPerson Then
            Call LoadBalance(strPerson): mblnNotBrush = False
            objItem.Selected = True: Exit Sub
        End If
    Next
    If txtChargePerson.Tag = txtChargePerson.Text Then
        '未过滤的,表示未找到
        mblnNotBrush = False
        Cancel = True: Exit Sub
    End If
    
    '肯定存在按姓名/编号/简码过虑的,所以需要选清空
    txtChargePerson.Text = "": txtChargePerson_LostFocus
   '收费员轧帐
    For Each objItem In lvwPerson_S.ListItems
        If objItem.Text = strPerson Then
            Call LoadBalance(strPerson): mblnNotBrush = False
            objItem.Selected = True: Exit Sub
        End If
    Next
    '未过滤的,表示未找到
    mblnNotBrush = False
    Cancel = True
End Sub

Private Sub picBalance_Resize()
    Err = 0: On Error Resume Next
    With picBalance
        lblBalance.Left = .ScaleLeft + 50
        lblBalance.Width = .ScaleWidth - .Left * 2
        lblBalance.Top = .ScaleTop + 50
        lblBalance.Height = .ScaleHeight - .Top * 2
    End With
End Sub

 Private Sub picPersonList_Resize()
    Err = 0: On Error Resume Next
    With picPersonList
        txtChargePerson.Top = .ScaleTop + 50
        lblChargePerson.Top = txtChargePerson.Top + (txtChargePerson.Height - lblChargePerson.Height) \ 2
        txtChargePerson.Width = .ScaleWidth - txtChargePerson.Left
        lvwPerson_S.Left = .ScaleLeft
        lvwPerson_S.Top = txtChargePerson.Top + txtChargePerson.Height + 50
        lvwPerson_S.Width = .ScaleWidth
        lvwPerson_S.Height = .ScaleHeight - lvwPerson_S.Top - 50
    End With
End Sub
 Private Sub picGroupList_Resize()
    Err = 0: On Error Resume Next
    With picGroupList
        lvwGroup_S.Left = .ScaleLeft
        lvwGroup_S.Top = .ScaleTop
        lvwGroup_S.Width = .ScaleWidth
        lvwGroup_S.Height = .ScaleHeight
    End With
End Sub

 Private Sub picOtherList_Resize()
    Err = 0: On Error Resume Next
    With picOtherList
        txtOtherPerson.Top = .ScaleTop + 50
        lblOtherPerson.Top = txtOtherPerson.Top + (txtOtherPerson.Height - lblOtherPerson.Height) \ 2
        txtOtherPerson.Width = .ScaleWidth - txtOtherPerson.Left
        lvwOther_S.Left = .ScaleLeft
        lvwOther_S.Top = txtOtherPerson.Top + txtOtherPerson.Height + 50
        lvwOther_S.Width = .ScaleWidth
        lvwOther_S.Height = .ScaleHeight - lvwOther_S.Top - 50
    End With
End Sub
Private Sub picPersonPage_Resize()
    Err = 0: On Error Resume Next
    With picPersonPage
        tbPage.Left = .ScaleLeft
        tbPage.Top = .ScaleTop
        tbPage.Width = .ScaleWidth
        tbPage.Height = .ScaleHeight
    End With
End Sub
Private Sub InitPage()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化页面控件
    '编制:刘兴洪
    '日期:2013-09-22 17:07:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, objItem As TabControlItem, objForm As Object
    Err = 0: On Error GoTo ErrHand:
    mblnNotBrush = True
    picPersonList.Visible = zlStr.IsHavePrivs(mstrPrivs, "收费员收款")
     If zlStr.IsHavePrivs(mstrPrivs, "收费员收款") Then
        Set objItem = tbPage.InsertItem(EM_PG_收费员, "收费员", picPersonList.hWnd, 0)
        objItem.Tag = EM_PG_收费员
    End If
    
    picGroupList.Visible = zlStr.IsHavePrivs(mstrPrivs, "财务组收款")
     If zlStr.IsHavePrivs(mstrPrivs, "财务组收款") Then
        Set objItem = tbPage.InsertItem(EM_PG_财务组, "财务组", picGroupList.hWnd, 0)
        objItem.Tag = EM_PG_财务组
     End If
    picOtherList.Visible = zlStr.IsHavePrivs(mstrPrivs, "其他人员收款")
    If zlStr.IsHavePrivs(mstrPrivs, "其他人员收款") Then
        Set objItem = tbPage.InsertItem(EM_PG_其他人员, "其他人员", picOtherList.hWnd, 0)
        objItem.Tag = EM_PG_其他人员
    End If
    
     With tbPage
        Set tbPage.PaintManager.Font = Me.Font
        .PaintManager.Position = xtpTabPositionBottom
        tbPage.Item(0).Selected = True
        .PaintManager.ClientFrame = xtpTabFrameSingleLine
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.StaticFrame = True
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutSizeToFit
    End With
    mblnNotBrush = False
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    mblnNotBrush = False
End Sub

Private Function InitPanel()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化区域控件
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-09-22 17:13:23
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPane As Pane, objMain As Pane, lngWidth As Long
    Dim lngBalanceHeight As Long
    
    lngWidth = 3435 \ Screen.TwipsPerPixelX
    lngBalanceHeight = 600 \ Screen.TwipsPerPixelY
    With dkpMan
        Set objMain = .CreatePane(mPaneIndex.EM_PN_收费员列表, lngWidth, 400, DockLeftOf, Nothing)
        objMain.Title = ""
        objMain.Options = PaneNoCloseable Or PaneNoCaption Or PaneNoFloatable Or PaneNoHideable
        objMain.Handle = picPersonPage.hWnd
        objMain.MinTrackSize.Width = Int(lngWidth * 0.5): objMain.MaxTrackSize.Width = lngWidth
        
        Set objPane = .CreatePane(mPaneIndex.EM_PN_暂存列表, 100, lngBalanceHeight, DockBottomOf, objMain)
        objPane.Title = "当前暂存金":
        objPane.Options = PaneNoCloseable Or PaneNoCaption Or PaneNoFloatable Or PaneNoHideable
        objPane.Handle = picBalance.hWnd
        objPane.MinTrackSize.Height = Int(lngBalanceHeight * 0.5): objPane.MaxTrackSize.Height = lngBalanceHeight * 1.5
        
        If zlStr.IsHavePrivs(mstrPrivs, "其他人员收款") Then
            Set mfrmPersonOther = New frmFinanceSupervisePersonOthers
            Load mfrmPersonOther
            Call mfrmPersonOther.zlInitVar(mlngModule, mstrPrivs)
        End If
        
        Set mfrmList = New frmFinaceSuperviseCollectList
        Load mfrmList
        Call mfrmList.zlInitVar(EM_TY_收费员, mlngModule, mstrPrivs)
        Set mobjDetailPane = .CreatePane(mPaneIndex.EM_PN_明细列表, 100, 100, DockRightOf, objMain)
        mobjDetailPane.Title = "":
        mobjDetailPane.Options = PaneNoCloseable Or PaneNoCaption Or PaneNoFloatable Or PaneNoHideable
        mobjDetailPane.Handle = mfrmList.hWnd
        
        .Options.ThemedFloatingFrames = True
        .Options.UseSplitterTracker = False '实时拖动
        .Options.AlphaDockingContext = True
        .Options.HideClient = True
    End With
End Function

Public Function zlRollingCurtainCollect(ByVal frmMain As Object, Optional blnCustomCollect As Boolean = False) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载收款
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-10-11 11:45:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str收费员 As String, lng人员ID As Long, lng缴款组ID As Long
    Dim strIDs As String
    Dim strSQL As String, rsTmp As ADODB.Recordset
    On Error GoTo errHandle
   Select Case Val(tbPage.Selected.Tag)
   Case EM_PG_收费员
        If lvwPerson_S.SelectedItem Is Nothing Then Exit Function
        str收费员 = lvwPerson_S.SelectedItem.Text
        lng人员ID = Val(lvwPerson_S.SelectedItem.Tag)
        lng缴款组ID = 0
        If blnCustomCollect Then
            '手功收款
            zlRollingCurtainCollect = frmFinaceSuperviseCustomInput.EditCard(frmMain, str收费员, lng人员ID, mlngModule, mstrPrivs)
            Exit Function
        End If
        strIDs = mfrmList.GetSelRollingCurtainIds
        If strIDs = "" Then
            MsgBox "未选中需要收款的轧帐记录", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
        zlRollingCurtainCollect = frmFinanceSuperviseRollingCurtainEdit.zlShowMe(frmMain, mlngModule, mstrPrivs, str收费员, lng人员ID, strIDs, lng缴款组ID)
   Case EM_PG_财务组
        If blnCustomCollect Then
            MsgBox "财务组不支持手工缴款操作!", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
        
        If lvwGroup_S.SelectedItem Is Nothing Then Exit Function
        
        str收费员 = lvwGroup_S.SelectedItem.SubItems(3)
        '76120,冉俊明,2014-8-5,没有权限“收费员收款”给财务组收款,点击轧帐收款按钮报错“未设置对象变量或 With block 变量”
        lng人员ID = Val(lvwGroup_S.SelectedItem.Tag)
        lng缴款组ID = Val(Mid(lvwGroup_S.SelectedItem.Key, 2))
        strIDs = mfrmList.GetSelRollingCurtainIds
        If strIDs = "" Then
            MsgBox "未选中需要收款的轧帐记录", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
        strSQL = "Select Distinct A.收款员,B.ID From 人员收缴记录 A,人员表 B Where A.收款员=B.姓名 And A.记录性质=3 And A.小组轧账ID In "
        strSQL = strSQL & " (Select Column_Value From Table(f_str2list([1]))) "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strIDs)
        If rsTmp.RecordCount > 1 Then
            MsgBox "当前轧帐的财务组收款记录存在多个组长,无法继续!", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        Else
            If Not rsTmp.EOF Then
                str收费员 = Nvl(rsTmp!收款员)
                lng人员ID = Val(Nvl(rsTmp!ID))
            End If
        End If
        zlRollingCurtainCollect = frmFinanceSuperviseRollingCurtainEdit.zlShowMe(frmMain, mlngModule, mstrPrivs, str收费员, lng人员ID, strIDs, lng缴款组ID)
    Case EM_PG_其他人员
        If lvwOther_S.SelectedItem Is Nothing Then Exit Function
        str收费员 = lvwOther_S.SelectedItem.Text
        lng人员ID = Val(lvwOther_S.SelectedItem.Tag)
        lng缴款组ID = 0
        If blnCustomCollect Then
            '手功收款
            zlRollingCurtainCollect = frmFinaceSuperviseCustomInput.EditCard(frmMain, str收费员, lng人员ID, mlngModule, mstrPrivs, True)
            Exit Function
        End If
        zlRollingCurtainCollect = mfrmPersonOther.SaveData()
    Case Else
        Exit Function
    End Select
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Property Get IsAllowCollect() As Boolean
  '是否允许收款
  Select Case Val(tbPage.Selected.Tag)
   Case EM_PG_收费员
        IsAllowCollect = mfrmList.IsSelRollingCurtainRecord
   Case EM_PG_财务组
        IsAllowCollect = mfrmList.IsSelRollingCurtainRecord
   Case EM_PG_其他人员
        IsAllowCollect = False
   Case Else
        Exit Property
    End Select
End Property
Public Property Get IsAllowOtherCollect() As Boolean
  '是否允许收款
  Select Case Val(tbPage.Selected.Tag)
   Case EM_PG_其他人员
        IsAllowOtherCollect = Not lvwOther_S.SelectedItem Is Nothing
   Case EM_PG_收费员, EM_PG_财务组
        IsAllowOtherCollect = False
   Case Else
        Exit Property
    End Select
End Property

Public Property Get IsAllowViewChargeList() As Boolean
  '是否允许查看明细
    Select Case Val(tbPage.Selected.Tag)
        Case EM_PG_收费员, EM_PG_财务组
            IsAllowViewChargeList = mfrmList.GetRollingCurtainID <> 0
        Case EM_PG_其他人员
            IsAllowViewChargeList = True
        Case Else
            Exit Property
    End Select
End Property
Public Property Get IsAllowCustomCollect() As Boolean
  '是否允许手功收款
  Select Case Val(tbPage.Selected.Tag)
   Case EM_PG_收费员
        IsAllowCustomCollect = Not lvwPerson_S.SelectedItem Is Nothing
   Case EM_PG_财务组
        IsAllowCustomCollect = False
   Case EM_PG_其他人员
        IsAllowCustomCollect = Not lvwOther_S.SelectedItem Is Nothing
   Case Else
        Exit Property
    End Select
End Property
Public Sub zlPrint(ByVal bytMode As Byte)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:输出列表信息
    '入参:bytMode=1-打印,2-预览,3-输出到Excel
    '编制:刘兴洪
    '日期:2013-09-13 10:23:30
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Select Case Val(tbPage.Selected.Tag)
    Case EM_PG_收费员
        Call mfrmList.zlPrint(bytMode)
   Case EM_PG_财务组
        Call mfrmList.zlPrint(bytMode)
   Case EM_PG_其他人员
        Call mfrmPersonOther.zlPrint(bytMode)
   Case Else: Exit Sub
   End Select
End Sub
Public Sub ShowChargeList(ByVal frmMain As Object)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示明细收款数据
    '编制:刘兴洪
    '日期:2013-09-16 17:33:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Select Case Val(tbPage.Selected.Tag)
    Case EM_PG_收费员
        Call mfrmList.ShowChargeList(frmMain)
   Case EM_PG_财务组
       Call mfrmList.ShowChargeList(frmMain)
   Case EM_PG_其他人员
       Call mfrmPersonOther.ShowChargeList(Me)
   Case Else: Exit Sub
   End Select
End Sub
Public Sub zlRefresh()
    Call LoadGroup: Call LoadPerson(False): Call LoadOtherPerson(False)
    Call LoadLocalePersonDetailData
End Sub

Public Sub CallCustomRpt(ByVal frmMain As Object, ByVal lngSys As Long, ByVal strRptCode As String)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:调用自定义报表
    '入参:lngSys-系统号
    '        strRptCode-报表编号
    '编制:刘兴洪
    '日期:2013-09-17 10:18:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Call mfrmList.CallCustomRpt(frmMain, lngSys, strRptCode)
End Sub
Public Property Get GetCashMoney() As Double
    '获取现金金额
    Select Case Val(tbPage.Selected.Tag)
    Case EM_PG_收费员, EM_PG_财务组
       GetCashMoney = mfrmList.GetCashMoney
    Case EM_PG_其他人员
        GetCashMoney = mfrmPersonOther.GetCashMoney
    Case Else
    End Select
End Property
Private Sub tbPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    '选择时,才提取Call LoadPerson(False)
    If mblnNotBrush Then Exit Sub
    mblnNotBrush = True
   Select Case Val(tbPage.Selected.Tag)
   Case EM_PG_收费员
        'If lvwPerson_S.Enabled And lvwPerson_S.Visible Then lvwPerson_S.SetFocus
        If lvwPerson_S.ListItems.Count <> 0 Then lvwPerson_S.SelectedItem.Selected = False
        mobjDetailPane.Handle = mfrmList.hWnd
        dkpMan.RecalcLayout
        mfrmPersonOther.Hide
   Case EM_PG_财务组
        'If lvwGroup_S.Enabled And lvwGroup_S.Visible Then lvwGroup_S.SetFocus
        If lvwGroup_S.ListItems.Count <> 0 Then lvwGroup_S.SelectedItem.Selected = False
        mobjDetailPane.Handle = mfrmList.hWnd
        dkpMan.RecalcLayout
        mfrmPersonOther.Hide
    Case EM_PG_其他人员
        'If lvwOther_S.Enabled And lvwOther_S.Visible Then lvwOther_S.SetFocus
        If lvwOther_S.ListItems.Count <> 0 Then lvwOther_S.SelectedItem.Selected = False
        mobjDetailPane.Handle = mfrmPersonOther.hWnd
        dkpMan.RecalcLayout
        mfrmList.Hide
   Case Else
        Exit Sub
   End Select
    mblnNotBrush = False
    lblBalance.Caption = ""
    mstrSelOther = ""
    'Call LoadLocalePersonDetailData
End Sub
Private Function LoadLocalePersonDetailData() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能: 加载指定人员的明细数据
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-09-26 11:29:34
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnBalance As Boolean, blnRollingCurtainMgr As Boolean
    Dim strPerson As String, lngGroupID As Long
    On Error GoTo errHandle
    '加载指定人员的轧帐记录
  Select Case Val(tbPage.Selected.Tag)
   Case EM_PG_收费员
        If Not lvwPerson_S.SelectedItem Is Nothing Then
             strPerson = lvwPerson_S.SelectedItem.Text
        End If
        
        '加载当前暂存金
        Call mfrmList.zlClearData
        blnBalance = LoadBalance(strPerson)
        blnRollingCurtainMgr = mfrmList.zlLoadCollectData(EM_TY_收费员, strPerson)
   Case EM_PG_财务组
        If Not lvwGroup_S.SelectedItem Is Nothing Then
             strPerson = lvwGroup_S.SelectedItem.SubItems(3)
             lngGroupID = Val(lvwGroup_S.SelectedItem.SubItems(1))
        End If
        '加载当前暂存金
        Call mfrmList.zlClearData
        blnBalance = LoadBalance(strPerson)
        blnRollingCurtainMgr = mfrmList.zlLoadCollectData(EM_TY_小组, strPerson, lngGroupID)
   Case EM_PG_其他人员
        If Not lvwOther_S.SelectedItem Is Nothing Then
             strPerson = lvwOther_S.SelectedItem.Text
        End If
        '加载当前暂存金
        blnBalance = LoadBalance(strPerson)
        mfrmPersonOther.zlLoadPersonData (strPerson)
   Case Else
        Exit Function
   End Select
    LoadLocalePersonDetailData = blnBalance Or blnRollingCurtainMgr
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub txtChargePerson_Change()
    If txtChargePerson.Text = txtChargePerson.Tag Then Exit Sub
    '进行过滤
    Call LoadPerson(True)
    If Not mblnNotBrush Then Call ClearData
End Sub
Private Sub txtChargePerson_GotFocus()
    If txtChargePerson.Text = txtChargePerson.Tag Then
        txtChargePerson.Text = ""
        txtChargePerson.ForeColor = lvwOther_S.ForeColor
    End If
    zlControl.TxtSelAll txtChargePerson
    zlCommFun.OpenIme False
End Sub

Private Sub txtChargePerson_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If lvwPerson_S.ListItems.Count = 1 Then
        lvwPerson_S.ListItems(1).Selected = True
        Call lvwPerson_S_ItemClick(lvwPerson_S.SelectedItem)
        If txtChargePerson.Enabled And txtChargePerson.Visible Then txtChargePerson.SetFocus
    End If
End Sub
Private Sub txtChargePerson_LostFocus()
    zlCommFun.OpenIme False
    If txtChargePerson.Text = "" Then
        txtChargePerson.ForeColor = &H80000000
        txtChargePerson.Text = "输入简码或汉字或编号"
    End If
End Sub
Private Function LoadBalance(ByVal strPerson As String) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载人员暂存金
    '入参:strPerson-人员
    '返回:加载成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-09-25 11:53:34
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    Dim strPreName As String
    
    On Error GoTo errHandle
    strSQL = "Select 收款员,结算方式,余额 From 人员缴款余额 where Instr(',' || [1] || ',' ,',' || 收款员 || ',') > 0 and 性质=1 Order By 收款员"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strPerson)
    With rsTemp
        strSQL = ""
        lblBalance.Caption = ""
        Do While Not .EOF
            If strPreName <> "" And strPreName <> !收款员 Then
                If lblBalance.Caption <> "" Then lblBalance.Caption = lblBalance.Caption & vbCrLf
                lblBalance.Caption = lblBalance.Caption & strPreName & "的暂存金:" & Mid(strSQL, 2)
                strSQL = ""
            End If
            strSQL = strSQL & " " & Nvl(!结算方式) & ":" & Format(Val(Nvl(!余额)), "0.00")
            strPreName = !收款员
            .MoveNext
        Loop
        If strSQL <> "" Then
            strSQL = Mid(strSQL, 2)
            If lblBalance.Caption <> "" Then lblBalance.Caption = lblBalance.Caption & vbCrLf
            lblBalance.Caption = lblBalance.Caption & strPreName & "的暂存金:" & strSQL
        End If
    End With
    LoadBalance = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Function
 
Private Sub txtOtherPerson_Change()
    If txtOtherPerson.Text = txtChargePerson.Tag Then Exit Sub
    '进行过滤
    Call LoadOtherPerson(True)
    If Not mblnNotBrush Then Call ClearData
End Sub
Private Sub txtOtherPerson_GotFocus()
    If txtOtherPerson.Text = txtOtherPerson.Tag Then
        txtOtherPerson.Text = ""
        txtOtherPerson.ForeColor = lvwOther_S.ForeColor
    End If
    zlControl.TxtSelAll txtOtherPerson
    zlCommFun.OpenIme False
End Sub

Private Sub txtOtherPerson_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If lvwOther_S.ListItems.Count <> 1 Then Exit Sub
    lvwOther_S.ListItems(1).Selected = True
    Call lvwOther_S_ItemClick(lvwOther_S.SelectedItem)
    If txtOtherPerson.Enabled And txtOtherPerson.Visible Then txtOtherPerson.SetFocus
End Sub

Private Sub txtOtherPerson_LostFocus()
    zlCommFun.OpenIme False
    If txtOtherPerson.Text = "" Then
        txtOtherPerson.ForeColor = &H80000000
        txtOtherPerson.Text = "输入简码或汉字或编号"
    End If
End Sub
Private Sub ShowPopup()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示弹出菜单
    '编制:刘兴洪
    '日期:2013-09-27 15:21:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPopup As CommandBarPopup
    Dim objCommandBar As CommandBar
    Dim objControl As CommandBarControl
    Set objCommandBar = mcbsMain.Add("PopupPati", xtpBarPopup)
    With objCommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_View_LargeICO, "大图标(&G)")
        Set objControl = .Add(xtpControlButton, conMenu_View_MinICO, "小图标(&M)")
        Set objControl = .Add(xtpControlButton, conMenu_View_ListICO, "列表(&L)")
        Set objControl = .Add(xtpControlButton, conMenu_View_DetailsICO, "详细资料(&D)")
  End With
  If Not objCommandBar Is Nothing Then objCommandBar.ShowPopup
End Sub
 
 '人员列表的显示方式
Public Property Get zlPersonListShowMode() As Integer
  Select Case Val(tbPage.Selected.Tag)
    Case EM_PG_收费员
        zlPersonListShowMode = lvwPerson_S.View
    Case EM_PG_财务组
        zlPersonListShowMode = lvwGroup_S.View
    Case EM_PG_其他人员
        zlPersonListShowMode = lvwOther_S.View
    End Select
End Property

Public Property Let zlPersonListShowMode(ByVal vNewValue As Integer)
   Select Case Val(tbPage.Selected.Tag)
    Case EM_PG_收费员
        lvwPerson_S.View = vNewValue
    Case EM_PG_财务组
        lvwGroup_S.View = vNewValue
    Case EM_PG_其他人员
        lvwOther_S.View = vNewValue
    End Select
End Property
Private Sub ClearData()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:清除相关数据
    '编制:刘兴洪
    '日期:2013-09-29 11:20:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
   Select Case Val(tbPage.Selected.Tag)
    Case EM_PG_收费员, EM_PG_财务组
          Call mfrmList.zlClearData
    Case EM_PG_其他人员
    End Select
End Sub
