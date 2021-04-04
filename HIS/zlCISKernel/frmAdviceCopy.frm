VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAdviceCopy 
   AutoRedraw      =   -1  'True
   Caption         =   "复制医嘱"
   ClientHeight    =   6540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9960
   Icon            =   "frmAdviceCopy.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   9960
   Begin VB.ComboBox cboQX 
      Height          =   300
      Left            =   3945
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Top             =   5820
      Width           =   1245
   End
   Begin VB.Frame fraLR 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5325
      Left            =   3270
      MousePointer    =   9  'Size W E
      TabIndex        =   16
      Top             =   870
      Width           =   45
   End
   Begin VSFlex8Ctl.VSFlexGrid vsPati 
      Height          =   5340
      Left            =   30
      TabIndex        =   7
      Top             =   840
      Width           =   3255
      _cx             =   5741
      _cy             =   9419
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
      BackColorSel    =   12632256
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   2000
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmAdviceCopy.frx":058A
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
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
      OwnerDraw       =   1
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   -1  'True
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
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   465
      TabIndex        =   13
      ToolTipText     =   "F1"
      Top             =   6045
      Width           =   1100
   End
   Begin VSFlex8Ctl.VSFlexGrid vsAdvice 
      Height          =   5340
      Left            =   3390
      TabIndex        =   8
      Top             =   840
      Width           =   6525
      _cx             =   11509
      _cy             =   9419
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
      BackColorSel    =   12632256
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   0
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   24
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   2000
      ColWidthMin     =   0
      ColWidthMax     =   5000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmAdviceCopy.frx":0652
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   1
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
      OwnerDraw       =   1
      Editable        =   2
      ShowComboButton =   1
      WordWrap        =   -1  'True
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
      FrozenCols      =   2
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   7935
      TabIndex        =   10
      Top             =   6045
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6825
      TabIndex        =   9
      Top             =   6045
      Width           =   1100
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "全清(&R)"
      Height          =   350
      Left            =   2745
      TabIndex        =   12
      ToolTipText     =   "Ctrl+R"
      Top             =   6045
      Width           =   1100
   End
   Begin VB.CommandButton cmdAll 
      Caption         =   "全选(&A)"
      Height          =   350
      Left            =   1635
      TabIndex        =   11
      ToolTipText     =   "Ctrl+A"
      Top             =   6045
      Width           =   1100
   End
   Begin VB.Frame fraPati 
      Height          =   900
      Left            =   45
      TabIndex        =   15
      Top             =   -75
      Width           =   9900
      Begin VB.ComboBox cboFinTim 
         Enabled         =   0   'False
         Height          =   300
         Left            =   3105
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   540
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.OptionButton optType 
         Caption         =   "完成就诊"
         Height          =   195
         Index           =   1
         Left            =   1230
         TabIndex        =   20
         Top             =   600
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.OptionButton optType 
         Caption         =   "正在就诊"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   19
         Top             =   600
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.PictureBox picDiag 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3255
         Picture         =   "frmAdviceCopy.frx":08EF
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   18
         Top             =   210
         Width           =   255
      End
      Begin VB.ComboBox cboBaby 
         Height          =   300
         Left            =   8370
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   195
         Width           =   1395
      End
      Begin VB.ComboBox cboTime 
         Height          =   300
         Left            =   3105
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   195
         Width           =   3870
      End
      Begin VB.CommandButton cmdPati 
         Height          =   240
         Left            =   1950
         Picture         =   "frmAdviceCopy.frx":7141
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "选择病人(F4)"
         Top             =   225
         Width           =   255
      End
      Begin VB.TextBox txtPati 
         Height          =   300
         Left            =   780
         TabIndex        =   1
         Top             =   195
         Width           =   1455
      End
      Begin VB.Label lblTim 
         AutoSize        =   -1  'True
         Caption         =   "就诊时间"
         Height          =   180
         Left            =   2340
         TabIndex        =   22
         Top             =   600
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label lblDiag 
         Caption         =   "病人诊断：   "
         Height          =   180
         Left            =   2400
         TabIndex        =   17
         Top             =   225
         Width           =   7335
      End
      Begin VB.Label lblBaby 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "婴儿(&B)"
         Height          =   180
         Left            =   7695
         TabIndex        =   5
         Top             =   255
         Width           =   630
      End
      Begin VB.Label lblTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "时间(&T)"
         Height          =   180
         Left            =   2430
         TabIndex        =   3
         Top             =   255
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "病人(&P)"
         Height          =   180
         Left            =   135
         TabIndex        =   0
         Top             =   255
         Width           =   630
      End
   End
   Begin MSComctlLib.ListView lvwPati 
      Height          =   3975
      Left            =   795
      TabIndex        =   14
      Top             =   450
      Visible         =   0   'False
      Width           =   5700
      _ExtentX        =   10054
      _ExtentY        =   7011
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "img32"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   12
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "病人"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "住院号"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "床号"
         Object.Width           =   1111
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "住院医师"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "性别"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "年龄"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "费别"
         Object.Width           =   1499
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "护理等级"
         Object.Width           =   2028
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   $"frmAdviceCopy.frx":7237
         Object.Width           =   2857
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "出院日期"
         Object.Width           =   2857
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "付款方式"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "病人类型"
         Object.Width           =   2646
      EndProperty
   End
End
Attribute VB_Name = "frmAdviceCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mfrmParent As Object
Private mMainPrivs As String
Private mbln护士站 As Boolean
Private mlng前提ID As Long
Private mlng病人ID As Long
Private mlng主页ID As Long
Private mlngPatiID As Long '外面传入的病人ID
Private mstrNoIn As String '外面传入的挂号单，即当前病人的挂号单
Private mstr挂号单 As String
Private mblnMoved As Boolean
Private mblnItem As Boolean
Private mstrIDs As String
Private mstrAlter As String
Private mlng婴儿 As String
Private mstr性别 As String

Private mtmp主页ID As Long
Private mtmp挂号单 As String
Private mlng病人科室id As Long
Private mlng婴儿科室ID As Long
Private mbln医技后续 As Boolean
Private mint来源 As Integer

Private Enum COL医嘱
    col选择 = 0
    col期效 = 1
    col时间 = 2
    col内容 = 3
    col总量 = 4
    col总量单位 = 5
    col单量 = 6
    col单量单位 = 7
    col频次 = 8
    col用法 = 9
    col嘱托 = 10
    col执行时间 = 11
    col执行科室 = 12
    colID = 13
    col相关ID = 14
    col诊疗类别 = 15
    col诊疗项目ID = 16
    col收费细目ID = 17
    col是否适用 = 18
    col毒理分类 = 19
    col价值分类 = 20
    col性别 = 21
    col单独应用 = 22
    col操作类型 = 23
    col项目名称 = 24
    col收费名称 = 25
End Enum

Private Enum mCtlID
    opt正在就诊 = 0
    opt完成就诊 = 1
End Enum

Private Const con项目撤档 = -1
Private Const con项目服务 = -2
Private Const con项目类别 = -3
Private Const con收费撤档 = -4
Private Const con收费服务 = -5

Private mbln麻醉类权限 As Boolean '没有下达麻醉类权限时为True,有下达麻醉类权限时为False
Private mbln毒性类权限 As Boolean '没有下达毒性类权限时为True,有下达毒性类权限时为False
Private mbln精神类权限 As Boolean '没有下达精神类权限时为True,有下达精神类权限时为False
Private mbln贵重类权限 As Boolean '没有下达贵重类权限时为True,有下达贵重类权限时为False
Private mintOutPreTime As Integer
Private mdtOutBegin As Date, mdtOutEnd As Date
Private mlngPre期效 As Long

Public Function ShowMe(ByVal frmParent As Object, ByVal strPrivs As String, lng病人ID As Long, varTime As Variant, blnMoved As Boolean, _
    Optional ByVal bln护士站 As Boolean, Optional ByVal lng前提ID As Long, Optional strAlter As String, Optional lng病人科室ID As Long, _
    Optional lng婴儿 As Long, Optional lng婴儿科室ID As Long, Optional str性别 As String) As String
'返回：lng病人ID,varTime=要复制医嘱的病人ID，主页ID(挂号单NO)
'      blnMoved=要复制病人的医嘱是否转出
'      strAlter=本次复制的医嘱中要切换期效的医嘱ID(组ID):123,456,...
'      ShowMe=要复制的医嘱的组ID串
    Set mfrmParent = frmParent
    mMainPrivs = strPrivs
    mbln护士站 = bln护士站
    mlng前提ID = lng前提ID
    mlng病人ID = lng病人ID
    mlng婴儿 = lng婴儿
    mlng病人科室id = lng病人科室ID
    mlng婴儿科室ID = lng婴儿科室ID
    mstr性别 = str性别
    If TypeName(varTime) = "String" Then
        mstr挂号单 = varTime
        mlng主页ID = 0
    Else
        mlng主页ID = varTime
        mstr挂号单 = ""
    End If
    mstrNoIn = mstr挂号单: mlngPatiID = mlng病人ID
    mblnMoved = blnMoved
    strAlter = "": mstrAlter = strAlter
    
    Me.Show 1, frmParent
    
    lng病人ID = mlng病人ID
    If TypeName(varTime) = "String" Then
        varTime = mstr挂号单
    Else
        varTime = mlng主页ID
    End If
    blnMoved = mblnMoved
    strAlter = mstrAlter
    ShowMe = mstrIDs
End Function

Private Function LoadPatients() As Boolean
'功能：读取与调用界面相同范围的病人列表
    Dim rsTmp As ADODB.Recordset
    Dim objItem As ListItem, strSQL As String
    Dim i As Integer, j As Integer
    Dim lng部门ID As Long, intBedLen As Long
    Dim curDate As Date, dtOutEnd As Date, dtOutBegin As Date
    Dim intTmp As Integer
    
    On Error GoTo errH
    
    '读取出院病人的时间范围
    curDate = zlDatabase.Currentdate
    intTmp = Val(zlDatabase.GetPara("出院病人结束间隔", glngSys, IIF(mbln护士站, p住院护士站, p住院医生站), 0))
    dtOutEnd = Format(curDate + intTmp, "yyyy-MM-dd 23:59:59")
    intTmp = Val(zlDatabase.GetPara("出院病人开始间隔", glngSys, IIF(mbln护士站, p住院护士站, p住院医生站), 1))
    dtOutBegin = Format(curDate - intTmp, "yyyy-MM-dd 00:00:00")
    
    If mlng前提ID <> 0 Then
        cmdPati.Visible = False
        If mstr挂号单 <> "" Then
            strSQL = "Select B.NO,B.病人ID,B.门诊号,B.姓名,B.性别,B.年龄,A.险类," & _
                " C.名称 as 科室,B.执行时间 as 就诊时间,Decode(B.执行状态,0,'等待就诊',1,'就诊完成',2,'正在就诊') as 就诊状态" & _
                " From 病人信息 A,病人挂号记录 B,部门表 C" & _
                " Where A.病人ID=B.病人ID And B.执行部门ID=C.ID And A.病人ID=[1] And B.NO=[2] And B.记录性质=1 And B.记录状态=1"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mstr挂号单)
        Else
            strSQL = _
                "Select A.病人ID,B.主页ID,B.住院号,NVL(B.姓名,A.姓名) 姓名,NVL(B.性别,A.性别) 性别 ,NVL(B.年龄,A.年龄) 年龄 ," & _
                " B.入院日期,B.出院日期,B.住院医师,B.出院病床 as 床号,B.费别," & _
                " B.险类,B.出院科室ID as 科室ID,B.当前病区ID as 病区ID,D.名称 as 科室,C.名称 as 护理等级," & _
                " B.状态,B.数据转出,Nvl(B.医疗付款方式,A.医疗付款方式) as 医疗付款方式,B.病人类型" & _
                " From 病人信息 A,病案主页 B,收费项目目录 C,部门表 D" & _
                " Where A.病人ID=B.病人ID And B.护理等级ID=C.ID(+) And B.出院科室ID=D.ID" & _
                " And A.病人ID=[1] And B.主页ID=[2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID)
        End If
    Else
        If mstrNoIn <> "" Then
            If optType(opt完成就诊).value Then
                strSQL = "Select B.NO,B.病人ID,B.门诊号,B.姓名,B.性别,B.年龄,A.险类,C.名称 as 科室,B.执行时间 as 就诊时间,'就诊完成' as 就诊状态" & _
                    " From 病人信息 A,病人挂号记录 B,部门表 C" & _
                    " Where A.病人ID=B.病人ID And B.执行部门ID=C.ID And B.执行状态+0=1 And B.执行人||''=[1] And B.记录性质=1 And B.记录状态=1" & _
                    " and B.执行时间 between [2] and [3] Order By NO"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.姓名, CDate(Format(mdtOutBegin, "yyyy-MM-dd 00:00:00")), CDate(Format(mdtOutEnd, "yyyy-MM-dd 23:59:59")))
                '过滤出病人后取第一病人
                If rsTmp.EOF Then
                    mstr挂号单 = ""
                    mlng病人ID = 0
                Else
                    mstr挂号单 = rsTmp!NO & ""
                    mlng病人ID = Val(rsTmp!病人ID & "")
                End If
            Else
                mstr挂号单 = mstrNoIn: mlng病人ID = mlngPatiID
                '提供当前医生正在就诊和最近已诊的病人清单供选择:因此这点暂不涉及判断和读取"H病人挂号记录"
                strSQL = "Select B.NO,B.病人ID,B.门诊号,B.姓名,B.性别,B.年龄,A.险类,Decode(B.执行状态,0,0,1,2,2,1) as 排序," & _
                    " C.名称 as 科室,B.执行时间 as 就诊时间,Decode(B.执行状态,0,'等待就诊',1,'就诊完成',2,'正在就诊') as 就诊状态" & _
                    " From 病人信息 A,病人挂号记录 B,部门表 C Where A.病人ID=B.病人ID And B.执行部门ID=C.ID And A.病人ID=[1] And B.NO=[2] And B.记录性质=1 And B.记录状态=1" & _
                    " Union " & _
                    " Select B.NO,B.病人ID,B.门诊号,B.姓名,B.性别,B.年龄,A.险类,1 as 排序,C.名称 as 科室,B.执行时间 as 就诊时间,'正在就诊' as 就诊状态" & _
                    " From 病人信息 A,病人挂号记录 B,部门表 C Where A.病人ID=B.病人ID And B.执行部门ID=C.ID And B.执行状态=2 And B.执行人||''=[3] And B.记录性质=1 And B.记录状态=1 Order By 排序,NO"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mstr挂号单, UserInfo.姓名)
            End If
        Else
            strSQL = "Select 出院科室ID as 科室ID,当前病区ID  as 病区ID,婴儿科室ID,婴儿病区ID From 病案主页 Where 病人ID=[1] And 主页ID=[2]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID)
                
            '提供当前科室/病区的在院病人清单供选择
            lng部门ID = IIF(mbln护士站, IIF(mlng婴儿科室ID <> 0, NVL(rsTmp!婴儿病区ID, 0), NVL(rsTmp!病区ID, 0)), IIF(mlng婴儿科室ID <> 0, NVL(rsTmp!婴儿科室ID, 0), NVL(rsTmp!科室ID, 0)))
            intBedLen = GetMaxBedLen(lng部门ID, Not mbln护士站)
            strSQL = _
                "Select decode(b.住院医师,[4],1,2) as 排序,A.病人ID,B.主页ID,B.住院号,NVL(B.姓名,A.姓名) 姓名,NVL(B.性别,A.性别) 性别 ,NVL(B.年龄,A.年龄) 年龄,B.入院日期,B.出院日期," & _
                " B.住院医师,LPAD(B.出院病床," & intBedLen & ",' ') as 床号,B.费别,B.险类," & _
                " B.出院科室ID as 科室ID,D.名称 as 科室,C.名称 as 护理等级,B.状态,B.数据转出," & _
                " Nvl(B.医疗付款方式,A.医疗付款方式) as 医疗付款方式,B.病人类型" & _
                " From 病人信息 A,病案主页 B,收费项目目录 C,部门表 D" & _
                " Where A.病人ID=B.病人ID And B.护理等级ID=C.ID(+) And B.出院科室ID=D.ID And A.病人ID=[1] And B.主页ID=[2]"
            strSQL = strSQL & " Union " & _
                "Select decode(b.住院医师,[4],1,2) as 排序,A.病人ID,B.主页ID,B.住院号,NVL(B.姓名,A.姓名) 姓名,NVL(B.性别,A.性别) 性别 ,NVL(B.年龄,A.年龄) 年龄,B.入院日期,B.出院日期," & _
                " B.住院医师,LPAD(B.出院病床," & intBedLen & ",' ') as 床号,B.费别,B.险类," & _
                " B.出院科室ID as 科室ID,D.名称 as 科室,C.名称 as 护理等级,B.状态,B.数据转出," & _
                " Nvl(B.医疗付款方式,A.医疗付款方式) as 医疗付款方式,B.病人类型" & _
                " From 病人信息 A,病案主页 B,收费项目目录 C,部门表 D,在院病人 R" & _
                " Where A.病人ID=B.病人ID And A.主页ID=B.主页ID And B.护理等级ID=C.ID(+)" & _
                " And B.出院科室ID=D.ID And a.病人ID=R.病人ID " & IIF(mbln护士站, " And A.当前病区ID=R.病区ID", "  And A.当前科室ID=R.科室ID") & _
                IIF(mbln护士站, " And (R.病区ID=[3] Or b.婴儿病区ID=[3])", " And (R.科室ID=[3] Or b.婴儿科室ID=[3])") & _
                IIF(Not mbln护士站 And InStr(mMainPrivs, "本科病人") = 0, " And B.住院医师=[4]", "")
            strSQL = strSQL & " Union " & _
                "Select decode(b.住院医师,[4],1,2) as 排序,A.病人ID,B.主页ID,B.住院号,NVL(B.姓名,A.姓名) 姓名,NVL(B.性别,A.性别) 性别 ,NVL(B.年龄,A.年龄) 年龄,B.入院日期,B.出院日期," & _
                " B.住院医师,LPAD(B.出院病床," & intBedLen & ",' ') as 床号,B.费别,B.险类," & _
                " B.出院科室ID as 科室ID,D.名称 as 科室,C.名称 as 护理等级,B.状态,B.数据转出," & _
                " Nvl(B.医疗付款方式,A.医疗付款方式) as 医疗付款方式,B.病人类型" & _
                " From 病人信息 A,病案主页 B,收费项目目录 C,部门表 D" & _
                " Where A.病人ID=B.病人ID And Nvl(B.主页ID,0)<>0 And B.护理等级ID=C.ID(+)" & _
                " And B.出院科室ID=D.ID And B.出院日期 between [5] and [6]" & _
                IIF(mbln护士站, " And B.当前病区ID+0=[3]", " And B.出院科室ID+0=[3]") & _
                IIF(Not mbln护士站 And InStr(mMainPrivs, "本科病人") = 0, " And B.住院医师=[4]", "") & _
                " Order by 排序,床号"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, mlng主页ID, lng部门ID, UserInfo.姓名, dtOutBegin, dtOutEnd)
        End If
    End If
    
    lvwPati.ListItems.Clear
    For i = 1 To rsTmp.RecordCount
        If mstr挂号单 <> "" Then
            Set objItem = lvwPati.ListItems.Add(, "_" & rsTmp!病人ID & "_" & rsTmp!NO, rsTmp!姓名, , "Pati")
            objItem.SubItems(1) = NVL(rsTmp!门诊号)
            objItem.SubItems(2) = NVL(rsTmp!性别)
            objItem.SubItems(3) = NVL(rsTmp!年龄)
            objItem.SubItems(4) = NVL(rsTmp!科室)
            objItem.SubItems(5) = Format(NVL(rsTmp!就诊时间), "yyyy-MM-dd HH:mm")
            objItem.SubItems(6) = NVL(rsTmp!就诊状态)
            objItem.SubItems(7) = NVL(rsTmp!NO)
            
            '保险病人用红色显示
            If Not IsNull(rsTmp!险类) Then
                Call SetItemColor(objItem, vbRed)
            End If
            
            '显示初始病人的信息
            If rsTmp!病人ID = mlng病人ID And rsTmp!NO = mstr挂号单 Then
                With objItem
                    txtPati.ForeColor = .ForeColor
                    txtPati.Text = .Text
                    .Selected = True '一定要选中当前病人
                End With
            End If
        Else
            Set objItem = lvwPati.ListItems.Add(, "_" & rsTmp!病人ID & "_" & rsTmp!主页ID, rsTmp!姓名, , "Pati")
            objItem.SubItems(1) = NVL(rsTmp!住院号)
            objItem.SubItems(2) = NVL(rsTmp!床号)
            objItem.SubItems(3) = NVL(rsTmp!住院医师)
            objItem.SubItems(4) = NVL(rsTmp!性别)
            objItem.SubItems(5) = NVL(rsTmp!年龄)
            objItem.SubItems(6) = NVL(rsTmp!科室)
            objItem.SubItems(7) = NVL(rsTmp!费别)
            objItem.SubItems(8) = NVL(rsTmp!护理等级)
            objItem.SubItems(9) = Format(rsTmp!入院日期, "yyyy-MM-dd HH:mm")
            objItem.SubItems(10) = Format(NVL(rsTmp!出院日期), "yyyy-MM-dd HH:mm")
            objItem.SubItems(11) = NVL(rsTmp!医疗付款方式)
            objItem.SubItems(12) = NVL(rsTmp!病人类型)
            objItem.Tag = NVL(rsTmp!数据转出, 0)
            
            '病人颜色
            Call SetItemColor(objItem, zlDatabase.GetPatiColor(NVL(rsTmp!病人类型)))
            
            '显示初始病人的信息
            If rsTmp!病人ID = mlng病人ID And rsTmp!主页ID = mlng主页ID Then
                With objItem
                    txtPati.ForeColor = .ForeColor
                    txtPati.Text = .Text
                    .Selected = True '一定要选中当前病人
                End With
            End If
        End If
        rsTmp.MoveNext
    Next
    
    LoadPatients = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub LoadPatiTime()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim i As Long
    
    On Error GoTo errH
    
    cboTime.Clear
    cboBaby.Clear
    vsPati.Rows = vsPati.FixedRows
    vsPati.Rows = vsPati.FixedRows + 1
    vsPati.Row = 1
    vsAdvice.Rows = vsAdvice.FixedRows
    vsAdvice.Rows = vsAdvice.FixedRows + 1
    vsAdvice.Row = 1
    
    If mstr挂号单 = "" And optType(opt正在就诊).Visible Then
        txtPati.Text = ""
        txtPati.Enabled = False
        cmdPati.Enabled = False
        MsgBox "未找到任何病人的就诊信息！", vbInformation, gstrSysName
        Exit Sub
    Else
        txtPati.Enabled = True
        cmdPati.Enabled = True
    End If
        
    If InStr(GetInsidePrivs(IIF(mstr挂号单 = "", p住院医嘱下达, p门诊医嘱下达)), ";复制他人医嘱;") = 0 Then
        txtPati.Locked = True
        cmdPati.Enabled = False
        If mstr挂号单 <> "" Then
            optType(opt正在就诊).Enabled = False
            optType(opt完成就诊).Enabled = False
        End If
    End If
    
    If mstr挂号单 <> "" Then
        strSQL = "Select A.ID,A.NO,A.发生时间,B.名称 as 科室,A.执行人 as 医生,A.诊室,A.急诊,A.复诊" & _
            " From 病人挂号记录 A,部门表 B Where A.执行部门ID=B.ID And A.病人ID=[1] And a.记录性质=1 And a.记录状态=1 Order by A.发生时间 Desc"
    Else
        strSQL = "Select A.主页ID,A.入院日期,B.名称 as 科室 From 病案主页 A,部门表 B" & _
            " Where A.出院科室ID=B.ID And A.病人ID=[1] And Nvl(A.主页ID,0)<>0 Order by A.主页ID Desc"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID)
    vsPati.Redraw = flexRDNone
    Do While Not rsTmp.EOF
        If mstr挂号单 <> "" Then
            cboTime.AddItem "[" & Format(rsTmp!发生时间, "yyyy-MM-dd HH:mm") & "]" & rsTmp!NO & "," & rsTmp!科室
            cboTime.ItemData(cboTime.NewIndex) = rsTmp!ID
            
            With vsPati
                If .RowData(.Rows - 1) <> 0 Then .AddItem ""
                .RowData(.Rows - 1) = Val(rsTmp!ID)
                .TextMatrix(.Rows - 1, 0) = Format(rsTmp!发生时间, "yyyy-MM-dd HH:mm")
                .TextMatrix(.Rows - 1, 1) = NVL(rsTmp!科室)
                .TextMatrix(.Rows - 1, 2) = NVL(rsTmp!医生)
                .TextMatrix(.Rows - 1, 3) = NVL(rsTmp!诊室)
                .TextMatrix(.Rows - 1, 4) = IIF(NVL(rsTmp!急诊) = 1, "", "")
                .TextMatrix(.Rows - 1, 5) = IIF(NVL(rsTmp!复诊) = 1, "", "")
                .TextMatrix(.Rows - 1, 6) = rsTmp!NO
            End With
            
            If rsTmp!NO = mstr挂号单 Then
                cboTime.ListIndex = cboTime.NewIndex
                vsPati.Row = vsPati.Rows - 1
            End If
        Else
            cboTime.AddItem "[" & Format(rsTmp!入院日期, "yyyy-MM-dd HH:mm") & "]第" & rsTmp!主页ID & "次住院," & rsTmp!科室
            cboTime.ItemData(cboTime.NewIndex) = rsTmp!主页ID
            If rsTmp!主页ID = mlng主页ID Then cboTime.ListIndex = cboTime.NewIndex
        End If
        rsTmp.MoveNext
    Loop
    
    With vsPati
        If vsPati.Tag = "加载部分" And .Rows > 4 Then
            For i = 4 To .Rows - 1
                .RowHidden(i) = True
            Next
            .AddItem ""
            .RowData(.Rows - 1) = -1
            .TextMatrix(.Rows - 1, 0) = "显示全部"
        End If
        .Row = .FixedRows
    End With
    Call vsPati.ShowCell(vsPati.Row, 0)
    vsPati.Redraw = flexRDDirect
    If cboTime.ListIndex = -1 Then cboTime.ListIndex = 0
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub SetItemColor(ByVal objItem As ListItem, ByVal lngColor As Long)
    Dim i As Long
    
    objItem.ForeColor = lngColor
    For i = 1 To objItem.ListSubItems.Count
        objItem.ListSubItems(i).ForeColor = lngColor
    Next
End Sub

Private Sub cboBaby_Click()
    Call LoadAdvice
End Sub

Private Sub cboFinTim_Click()
'功能：当时间范围是指定是，弹出时间选择窗体
    Dim intDateCount As Integer
    Dim datCurr As Date
    
    With cboFinTim
        intDateCount = .ItemData(.ListIndex)
        If .ListIndex = mintOutPreTime And intDateCount <> -1 Then Exit Sub
        If intDateCount = -1 Then
            If Not frmSelectTime.ShowMe(Me, mdtOutBegin, mdtOutEnd, cboFinTim) Then
                '取消时恢复原来的选择
                Call cbo.SetIndex(.hwnd, mintOutPreTime)
                Exit Sub
            End If
        Else
            mdtOutEnd = CDate(.Tag)
            mdtOutBegin = mdtOutEnd - intDateCount
        End If

        .ToolTipText = "范围：" & Format(mdtOutBegin, "yyyy-MM-dd 00:00") & " 至 " & Format(mdtOutEnd, "yyyy-MM-dd 23:59")
        lblTim.ToolTipText = .ToolTipText
        mintOutPreTime = .ListIndex
    End With
    datCurr = CDate(cboFinTim.Tag)
    
    Call LoadPatients
    Call LoadPatiTime
End Sub

Private Sub cboTime_Click()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, blnVisible As Boolean
    
    If cboTime.ListIndex = -1 Then Exit Sub
    
    On Error GoTo errH
    
    cboBaby.Clear
    vsAdvice.Rows = vsAdvice.FixedRows
    vsAdvice.Rows = vsAdvice.FixedRows + 1
    vsAdvice.Row = 1
    
    If mstr挂号单 <> "" Then
        strSQL = "Select Distinct A.婴儿 From 病人医嘱记录 A,病人挂号记录 B Where A.挂号单=B.NO And B.ID=[2] Order by Nvl(A.婴儿,0)"
    Else
        strSQL = "Select Distinct 婴儿 From 病人医嘱记录 Where 病人ID=[1] And 主页ID=[2] Order by Nvl(婴儿,0)"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, cboTime.ItemData(cboTime.ListIndex))
    Do While Not rsTmp.EOF
        If NVL(rsTmp!婴儿, 0) = 0 Then
            cboBaby.AddItem "病人医嘱"
        Else
            cboBaby.AddItem "婴儿 " & rsTmp!婴儿 & " 医嘱"
        End If
        cboBaby.ItemData(cboBaby.NewIndex) = NVL(rsTmp!婴儿, 0)
        If NVL(rsTmp!婴儿, 0) = mlng婴儿 Then cboBaby.ListIndex = cboBaby.NewIndex
        rsTmp.MoveNext
    Loop
    If cboBaby.ListIndex = -1 And cboBaby.ListCount > 0 Then cboBaby.ListIndex = 0
    Call LoadDiag
    
    blnVisible = cboBaby.ListCount > 0
    If cboBaby.ListCount = 1 Then
        If cboBaby.ItemData(cboBaby.ListIndex) = 0 Then
            blnVisible = False
        End If
    End If
    cboBaby.Visible = blnVisible
    lblBaby.Visible = blnVisible
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub LoadDiag()
'功能，加载该次就诊的诊断信息
    Dim lng就诊ID As Long
    Dim strSQL As String, rsTmp As Recordset
    Dim i As Integer
    Dim strDiag As String
    On Error GoTo errH
    
    lng就诊ID = cboTime.ItemData(cboTime.ListIndex)
    strSQL = "Select 诊断类型,诊断描述 From 病人诊断记录 Where 主页id = [2] And 病人id = [1] And 诊断类型 in (11,1) Order By 诊断次序"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, lng就诊ID)
    picDiag.Tag = ""
    lblDiag.Caption = "病人诊断：   "
    If rsTmp.RecordCount > 0 Then
        For i = 1 To rsTmp.RecordCount
            If 1 = Val(rsTmp!诊断类型) Then
                strDiag = strDiag & "," & rsTmp!诊断描述
                picDiag.Tag = picDiag.Tag & "【西】" & rsTmp!诊断描述 & vbCrLf
            Else
                picDiag.Tag = picDiag.Tag & "【中】" & rsTmp!诊断描述 & vbCrLf
            End If
            rsTmp.MoveNext
        Next
        lblDiag.Caption = lblDiag.Caption & Mid(strDiag, 2)
    End If
    If picDiag.Tag = "" Then
        picDiag.Tag = "没有任何诊断！"
    Else '去掉末尾的回车符
        picDiag.Tag = Mid(picDiag.Tag, 1, Len(picDiag.Tag) - 2)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdALL_Click()
    Dim i As Long, lngEnd As Long
    
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            If CheckCanSelGroup(i, False) Then
                Call SelGroup(i, 1, lngEnd)
            End If
            If i < lngEnd Then i = lngEnd
        Next
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdClear_Click()
    Dim i As Long
    
    With vsAdvice
        For i = .FixedRows To .Rows - 1
            .TextMatrix(i, col选择) = 0
        Next
    End With
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub cmdOK_Click()
    Dim lngID As Long, i As Long
    Dim strIDs As String, strAlter As String
    
    With vsAdvice
        '取一组医嘱的ID
        For i = .FixedRows To .Rows - 1
            If Val(.TextMatrix(i, col选择)) <> 0 Then
                lngID = Val(.TextMatrix(i, colID))
                If lngID <> 0 Then

                    '选择复制部份
                    If InStr(strIDs & ",", "," & lngID & ",") = 0 Then
                        strIDs = strIDs & "," & lngID
                    End If
                    
                    '切换期效部份
                    If .TextMatrix(i, col期效) <> .Cell(flexcpData, i, col期效) Then
                        If InStr(strAlter & ",", "," & lngID & ",") = 0 Then
                            strAlter = strAlter & "," & lngID
                        End If
                    End If
                End If
            End If
        Next
        strAlter = Mid(strAlter, 2)
        strIDs = Mid(strIDs, 2)
        If strIDs = "" Then
            MsgBox "请选择要复制的医嘱。", vbInformation, gstrSysName
            Exit Sub
        End If
    End With
    
    mstrAlter = strAlter
    mstrIDs = strIDs
    mstr挂号单 = mtmp挂号单
    mlng主页ID = mtmp主页ID
    
    Unload Me
End Sub

Private Sub cmdPati_Click()
    If mstr挂号单 <> "" Then
        lvwPati.ListItems("_" & mlng病人ID & "_" & mstr挂号单).Selected = True
    Else
        lvwPati.ListItems("_" & mlng病人ID & "_" & mlng主页ID).Selected = True
    End If
    lvwPati.SelectedItem.EnsureVisible
    lvwPati.Left = txtPati.Left + fraPati.Left
    lvwPati.Top = txtPati.Top + txtPati.Height + fraPati.Top
    lvwPati.Height = vsAdvice.Height - 300
    lvwPati.ZOrder
    lvwPati.Visible = True
    lvwPati.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        Call cmdHelp_Click
    ElseIf KeyCode = vbKeyEscape Then
        If lvwPati.Visible Then
            lvwPati.Visible = False
        Else
            Unload Me
        End If
    ElseIf KeyCode = vbKeyF4 Or KeyCode = vbKeyDown Then
        If Not (KeyCode = vbKeyDown And Shift <> vbAltMask) Then
            If Me.ActiveControl Is txtPati Then
                If cmdPati.Visible And cmdPati.Enabled Then cmdPati_Click
            End If
        End If
    ElseIf KeyCode = vbKeyA And Shift = vbCtrlMask Then
        Call cmdALL_Click
    ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
        Call cmdClear_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub Form_Load()
    Dim strLvw As String
    
    mbln医技后续 = Val(zlDatabase.GetPara("医技医嘱后续处理", glngSys, p住院医嘱发送)) <> 0
    lvwPati.SmallIcons = frmIcons.imgPati
    If mstr挂号单 <> "" Then
        strLvw = "病人,1000,0,1;门诊号,1000,0,1;性别,600,0,1;年龄,600,0,1;科室,1000,0,1;就诊时间,1620,0,1;就诊状态,1000,2,1;挂号单,1000,0,1"
    Else
        strLvw = "病人,1000,0,1;住院号,1000,0,1;床号,630,0,1;住院医师,1000,0,1;性别,600,0,1;年龄,600,0,1;科室,1000,1,0;费别,850,0,1;护理等级,1150,0,1;入院日期,1620,0,1;出院日期,1620,0,1;付款方式,1500,0,1;病人类型,1500,0,1"
    End If
    Call InitAdviceTable
    Call zlControl.LvwSelectColumns(lvwPati, strLvw, True)
    Call RestoreWinState(Me, App.ProductName, IIF(mstr挂号单 <> "", 1, 2))
    If mlng主页ID <> 0 Then
        vsAdvice.FrozenCols = col期效 + 1
    Else
        vsAdvice.FrozenCols = col选择 + 1
    End If
    
    If mstr挂号单 <> "" And mlng前提ID = 0 Then
        Call InitSelectTime
    End If
    
    
    '列表选择就诊记录方式，只有门诊可用
    If mlng前提ID <> 0 Or mstr挂号单 = "" And mlng主页ID <> 0 Then
        vsPati.Visible = False
        fraLR.Visible = False
        lblDiag.Visible = False
        picDiag.Visible = False
        mbln麻醉类权限 = InStr(GetTsPrivs(p住院医嘱下达), ";下达麻醉药嘱;") = 0
        mbln毒性类权限 = InStr(GetTsPrivs(p住院医嘱下达), ";下达毒性药嘱;") = 0
        mbln精神类权限 = InStr(GetTsPrivs(p住院医嘱下达), ";下达精神药嘱;") = 0
        mbln贵重类权限 = InStr(GetTsPrivs(p住院医嘱下达), ";下达贵重药嘱;") = 0
        mint来源 = 2
        cboQX.Visible = True
        cboQX.Clear
        cboQX.AddItem "所有"
        cboQX.AddItem "长嘱"
        cboQX.AddItem "临嘱"
        cboQX.ListIndex = 0
        mlngPre期效 = -1
    Else
        lblTime.Visible = False
        cboTime.Visible = False
        mbln麻醉类权限 = InStr(GetTsPrivs(p门诊医嘱下达), ";下达麻醉药嘱;") = 0
        mbln毒性类权限 = InStr(GetTsPrivs(p门诊医嘱下达), ";下达毒性药嘱;") = 0
        mbln精神类权限 = InStr(GetTsPrivs(p门诊医嘱下达), ";下达精神药嘱;") = 0
        mbln贵重类权限 = InStr(GetTsPrivs(p门诊医嘱下达), ";下达贵重药嘱;") = 0
        mint来源 = 1
        cboQX.Visible = False
    End If
    Call LoadPatients
    vsPati.Tag = "加载部分"
    Call LoadPatiTime
    mstrIDs = ""
End Sub

Private Sub cboQX_Click()
'功能：过滤医嘱
    If Not Me.Visible Then Exit Sub
    If cboQX.ListIndex <> mlngPre期效 Then
        mlngPre期效 = cboQX.ListIndex
        Call LoadAdvice
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    fraPati.Top = -75
    fraPati.Left = 0
    fraPati.Width = Me.ScaleWidth
    If mstr挂号单 <> "" And mlng前提ID = 0 Then
        fraPati.Height = 900
        optType(opt正在就诊).Visible = True
        optType(opt完成就诊).Visible = True
    Else
        fraPati.Height = 600
    End If
    
    If fraPati.Width - cboBaby.Width - 200 > 7500 Then
        cboBaby.Left = fraPati.Width - cboBaby.Width - 150
        lblBaby.Left = cboBaby.Left - lblBaby.Width - 30
    End If
    
    picDiag.Top = lblDiag.Top - 40
    picDiag.Left = lblDiag.Left + 900
    
    vsPati.Left = 0
    vsPati.Top = fraPati.Top + fraPati.Height
    vsPati.Height = Me.ScaleHeight - vsAdvice.Top - cmdOK.Height * 1.6
    
    fraLR.Left = vsPati.Width
    fraLR.Top = vsPati.Top
    fraLR.Height = vsPati.Height
    
    vsAdvice.Left = IIF(vsPati.Visible, vsPati.Width + fraLR.Width, 0)
    vsAdvice.Top = fraPati.Top + fraPati.Height
    vsAdvice.Width = Me.ScaleWidth - IIF(vsPati.Visible, vsPati.Width + fraLR.Width, 0)
    vsAdvice.Height = vsPati.Height
        
    cmdHelp.Top = Me.ScaleHeight - cmdAll.Height * 1.3
    cmdAll.Top = cmdHelp.Top
    cmdClear.Top = cmdAll.Top
    cmdOK.Top = cmdAll.Top
    cmdCancel.Top = cmdAll.Top
    
    If Me.ScaleWidth - cmdCancel.Width - (cmdHelp.Left + cmdHelp.Width / 3) < 5000 Then
        cmdCancel.Left = 5000
    Else
        cmdCancel.Left = Me.ScaleWidth - cmdCancel.Width - (cmdHelp.Left + cmdHelp.Width / 3)
    End If
    cmdOK.Left = cmdCancel.Left - cmdOK.Width
    
    cboQX.Left = cmdHelp.Left
    cboQX.Top = cmdHelp.Top - 350
    cboQX.Width = cmdHelp.Width
    Me.Refresh
End Sub
 
Private Sub optType_Click(Index As Integer)

    cboFinTim.Enabled = optType(opt完成就诊).value
    If Not (InStr(";" & mMainPrivs & ";", ";参数设置;") > 0 And cboFinTim.Enabled) Then cboFinTim.Enabled = False
    
    lblTim.Visible = optType(opt完成就诊).value
    cboFinTim.Visible = lblTim.Visible
    
    Call LoadPatients
    Call LoadPatiTime
End Sub

Private Sub InitSelectTime()
    Dim datCurr As Date, intStart As Integer, intDay As Integer
    Dim blnSetPar As Boolean
    
    With cboFinTim
        .Clear
        .AddItem "一周内"
        .ItemData(.NewIndex) = 7
        .AddItem "15天内"
        .ItemData(.NewIndex) = 15
        .AddItem "30天内"
        .ItemData(.NewIndex) = 30
        .AddItem "60天内"
        .ItemData(.NewIndex) = 60
        .AddItem "[指定...]"
        .ItemData(.NewIndex) = -1
        
        .Tag = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    End With
    
    datCurr = CDate(cboFinTim.Tag)

    blnSetPar = InStr(";" & mMainPrivs & ";", ";参数设置;") > 0
    intStart = Val(zlDatabase.GetPara("已诊病人结束间隔", glngSys, p门诊医生站, "0", Array(lblTim, cboFinTim), blnSetPar))
    intDay = Val(zlDatabase.GetPara("已诊病人开始间隔", glngSys, p门诊医生站, "7", Array(lblTim, cboFinTim), blnSetPar))
    
    mdtOutEnd = Format(datCurr + intStart, "yyyy-MM-dd 23:59:59")
    mdtOutBegin = Format(mdtOutEnd - intDay, "yyyy-MM-dd 00:00:00")
     
    cboFinTim.ToolTipText = Format(mdtOutBegin, "yyyy-MM-dd  00:00") & " - " & Format(mdtOutEnd, "yyyy-MM-dd 23:59")
    lblTim.ToolTipText = cboFinTim.ToolTipText
    
    If intStart = 0 Then
        Select Case intDay
        Case 7
            mintOutPreTime = 0
        Case 15
            mintOutPreTime = 1
        Case 30
            mintOutPreTime = 2
        Case 60
            mintOutPreTime = 3
        Case Else
            mintOutPreTime = 4
        End Select
    Else
        mintOutPreTime = 4
    End If
    
    Call cbo.SetIndex(cboFinTim.hwnd, mintOutPreTime)
End Sub

Private Sub picDiag_DblClick()
    MsgBox picDiag.Tag, vbInformation, Me.Caption
End Sub

Private Sub picDiag_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    zlCommFun.ShowTipInfo picDiag.hwnd, picDiag.Tag, True
    If X >= 0 And X <= picDiag.Width And Y >= 0 And Y <= picDiag.Height Then
        SetCapture picDiag.hwnd
    Else
        ReleaseCapture
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName, IIF(mstr挂号单 <> "", 1, 2))
End Sub

Private Sub fraLR_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If vsPati.Width + X < 1000 Or vsAdvice.Width - X < 1000 Then Exit Sub
        fraLR.Left = fraLR.Left + X
        vsPati.Width = vsPati.Width + X
        vsAdvice.Left = vsAdvice.Left + X
        vsAdvice.Width = vsAdvice.Width - X
        
        Me.Refresh
    End If
End Sub

Private Sub lvwPati_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Call zlControl.LvwSortColumn(lvwPati, ColumnHeader.Index)
End Sub

Private Sub lvwPati_DblClick()
    If mblnItem Then Call lvwPati_KeyPress(13)
End Sub

Private Sub lvwPati_ItemClick(ByVal Item As MSComctlLib.ListItem)
    mblnItem = True
End Sub

Private Sub lvwPati_KeyPress(KeyAscii As Integer)
    Dim lng病人ID As Long, lng主页ID As Long, strNO As String
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Not lvwPati.SelectedItem Is Nothing Then
            lng病人ID = Val(Split(Mid(lvwPati.SelectedItem.Key, 2), "_")(0))
            If mstr挂号单 <> "" Then
                strNO = Split(Mid(lvwPati.SelectedItem.Key, 2), "_")(1)
                If lng病人ID = mlng病人ID And strNO = mstr挂号单 Then
                    lvwPati.Visible = False
                    vsAdvice.SetFocus: Exit Sub
                End If
                With lvwPati.SelectedItem
                    mlng病人ID = lng病人ID
                    mstr挂号单 = strNO
                    mblnMoved = zlDatabase.NOMoved("病人挂号记录", strNO)
                    
                    txtPati.Text = .Text
                    txtPati.ForeColor = .ForeColor
                End With
            Else
                lng主页ID = Val(Split(Mid(lvwPati.SelectedItem.Key, 2), "_")(1))
                If lng病人ID = mlng病人ID And lng主页ID = mlng主页ID Then
                    lvwPati.Visible = False
                    vsAdvice.SetFocus: Exit Sub
                End If
                With lvwPati.SelectedItem
                    mlng病人ID = lng病人ID
                    mlng主页ID = lng主页ID
                    mblnMoved = Val(.Tag) = 1
                    
                    txtPati.Text = .Text
                    txtPati.ForeColor = .ForeColor
                End With
            End If
            lvwPati.Visible = False
            
            vsPati.Tag = "加载部分"
            
            Call LoadPatiTime
            
            If vsPati.Visible Then
                vsPati.SetFocus
            Else
                vsAdvice.SetFocus
            End If
        End If
    End If
End Sub

Private Sub lvwPati_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnItem = False
End Sub

Private Sub lvwPati_Validate(Cancel As Boolean)
    lvwPati.Visible = False
End Sub

Private Sub txtPati_GotFocus()
    Call zlControl.TxtSelAll(txtPati)
End Sub

Private Function RowIn一并给药(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long) As Boolean
'功能：判断指定行是否在一并给药的范围中,如果是,同时返回行号范围
    Dim i As Long, blnTmp As Boolean
    With vsAdvice
        If .TextMatrix(lngRow, col诊疗类别) = "" Then Exit Function
        If InStr(",5,6,", .TextMatrix(lngRow, col诊疗类别)) = 0 Then Exit Function
        If Val(.TextMatrix(lngRow - 1, col相关ID)) = Val(.TextMatrix(lngRow, col相关ID)) Then
            blnTmp = True
        ElseIf lngRow + 1 <= .Rows - 1 Then
            If Val(.TextMatrix(lngRow + 1, col相关ID)) = Val(.TextMatrix(lngRow, col相关ID)) Then
                blnTmp = True
            End If
        End If
        If blnTmp Then
            lngBegin = lngRow
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, col相关ID)) = Val(.TextMatrix(lngRow, col相关ID)) Then
                    lngBegin = i
                Else
                    Exit For
                End If
            Next
            lngEnd = lngRow
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, col相关ID)) = Val(.TextMatrix(lngRow, col相关ID)) Then
                    lngEnd = i
                Else
                    Exit For
                End If
            Next
        End If
        RowIn一并给药 = blnTmp
    End With
End Function

Private Function LoadAdvice() As Boolean
'功能：读取当前病人指定的医嘱
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long, j As Long
    Dim lng就诊ID As Long, int婴儿 As Integer
    Dim strDepartments As String
    Dim strWhere期效 As String
    
    If cboTime.ListIndex = -1 Then Exit Function
    If cboBaby.ListIndex = -1 Then Exit Function
    lng就诊ID = cboTime.ItemData(cboTime.ListIndex)
    int婴儿 = cboBaby.ItemData(cboBaby.ListIndex)
    
    If cboQX.Visible Then
        Select Case cboQX.ListIndex
        Case 1
            strWhere期效 = " and a.医嘱期效=0"
        Case 2
            strWhere期效 = " and a.医嘱期效=1"
        End Select
    End If
    
    On Error GoTo errH
    
    '排开撤档和不服务于的内容
    strSQL = "Select Distinct (Select Count(1) From 诊疗适用科室 Where 项目ID=b.ID) as 适用科室数,g.科室id as 适用科室ID,A.ID,A.序号,A.相关ID,A.医嘱期效,A.开始执行时间,A.诊疗项目ID," & _
        " A.医嘱内容,A.执行性质,A.单次用量,A.执行频次,A.医生嘱托,B.单独应用,B.操作类型,Nvl(C.名称,Decode(Nvl(A.执行性质,0),5,'-')) as 执行科室 ,A.执行时间方案,A.收费细目ID," & _
        " A.标本部位,A.检查方法,B.类别,a.诊疗类别,B.名称,B.计算单位,A.总给予量 as 总量,E.门诊包装,E.门诊单位,E.住院包装,E.住院单位," & _
        " D.计算单位 as 散装单位,B.撤档时间,B.服务对象,D.撤档时间 as 收费撤档,D.服务对象 as 收费服务,h.毒理分类,h.价值分类,Nvl(b.适用性别, 0) As 性别,b.名称 as 项目名称,d.名称 as 收费名称"
    If mstr挂号单 <> "" Then
        strSQL = strSQL & ",A.挂号单" & _
            " From 病人医嘱记录 A,诊疗项目目录 B,部门表 C,收费项目目录 D,药品规格 E,病人挂号记录 R,诊疗适用科室 G,药品特性 H" & _
            " Where A.诊疗项目ID=B.ID(+) And A.执行科室ID=C.ID(+) And A.收费细目ID=D.ID(+) And b.id=g.项目id(+) And h.药名ID(+)=b.ID And g.科室ID(+)=[4] And A.收费细目ID=E.药品ID(+)" & _
            " And Nvl(A.执行标记,0)<>-1 And A.医嘱状态 Not IN(2,4) And A.开始执行时间 is Not Null And Nvl(A.医嘱状态,0)<>-1 And A.病人来源=1" & _
            " And A.病人ID+0=[1] And A.挂号单=R.NO And R.ID=[2] And Nvl(A.婴儿,0)=[3]" & strWhere期效 & _
            " Order by A.序号"
    Else
        strSQL = strSQL & _
            " From 病人医嘱记录 A,诊疗项目目录 B,部门表 C,收费项目目录 D,药品规格 E,诊疗适用科室 G,药品特性 H" & _
            " Where A.诊疗项目ID=B.ID(+) And A.执行科室ID=C.ID(+) And A.收费细目ID=D.ID(+)  And b.id=g.项目id(+) And h.药名ID(+)=b.ID And g.科室ID(+)=[4] And A.收费细目ID=E.药品ID(+)" & _
            " And Nvl(A.执行标记,0)<>-1 And A.医嘱状态 Not IN(2,4) And A.开始执行时间 is Not Null And Nvl(A.医嘱状态,0)<>-1 And A.病人来源=2" & _
            " And A.病人ID=[1] And A.主页ID=[2] And Nvl(A.婴儿,0)=[3]" & strWhere期效 & _
            " Order by A.序号"
    End If
    If mblnMoved Then
        strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, lng就诊ID, int婴儿, mlng病人科室id)
    With vsAdvice
        .Redraw = flexRDNone
        .Rows = .FixedRows '清除表格内容
        If rsTmp.EOF Then
            .Rows = .FixedRows + 1
        Else
            .Rows = .FixedRows + rsTmp.RecordCount
            For i = 1 To rsTmp.RecordCount
                If i = 1 Then
                    If mstr挂号单 <> "" Then
                        mtmp挂号单 = rsTmp!挂号单
                    Else
                        mtmp主页ID = lng就诊ID
                    End If
                End If
            
                .TextMatrix(i, col选择) = 0
                .TextMatrix(i, colID) = rsTmp!ID
                .TextMatrix(i, col相关ID) = NVL(rsTmp!相关ID)
                .TextMatrix(i, col诊疗类别) = NVL(rsTmp!类别, "*")
                .TextMatrix(i, col诊疗项目ID) = NVL(rsTmp!诊疗项目ID)
                .TextMatrix(i, col收费细目ID) = NVL(rsTmp!收费细目ID)
                .TextMatrix(i, col期效) = IIF(NVL(rsTmp!医嘱期效, 0) = 0, "长嘱", "临嘱")
                .Cell(flexcpData, i, col期效) = .TextMatrix(i, col期效)
                .TextMatrix(i, col时间) = Format(rsTmp!开始执行时间, "yyyy-MM-dd HH:mm")
                .TextMatrix(i, col内容) = rsTmp!医嘱内容
                If .TextMatrix(i, col诊疗类别) = "D" And NVL(rsTmp!标本部位) <> "" Then
                    .Cell(flexcpData, i, col内容) = "[" & NVL(rsTmp!标本部位) & "]" & NVL(rsTmp!检查方法) '检验标本
                Else
                    .Cell(flexcpData, i, col内容) = NVL(rsTmp!标本部位)
                End If
                .TextMatrix(i, col单量) = FormatEx(NVL(rsTmp!单次用量), 4)
                .Cell(flexcpData, i, col单量) = NVL(rsTmp!名称)
                If Val(rsTmp!适用科室数 & "") > 0 And rsTmp!适用科室ID & "" = "" Then
                    .TextMatrix(i, col是否适用) = "1"
                End If
                .TextMatrix(i, col毒理分类) = NVL(rsTmp!毒理分类)
                .TextMatrix(i, col价值分类) = NVL(rsTmp!价值分类)
                .TextMatrix(i, col性别) = Decode(Val(rsTmp!性别), 0, "未知", 1, "男", 2, "女")
                If Not IsNull(rsTmp!单次用量) Then
                    If NVL(rsTmp!类别) = "4" Then
                        .TextMatrix(i, col单量单位) = NVL(rsTmp!散装单位)
                    Else
                        .TextMatrix(i, col单量单位) = NVL(rsTmp!计算单位)
                    End If
                End If
                If InStr(",5,6,", NVL(rsTmp!类别, "*")) > 0 Then
                    If mstr挂号单 <> "" Then
                        If Not IsNull(rsTmp!总量) And Not IsNull(rsTmp!门诊包装) Then
                            .TextMatrix(i, col总量) = FormatEx(rsTmp!总量 / rsTmp!门诊包装, 5)
                        End If
                        If NVL(rsTmp!医嘱期效, 0) = 1 Then
                            .TextMatrix(i, col总量单位) = NVL(rsTmp!门诊单位)
                        End If
                    Else
                        If Not IsNull(rsTmp!总量) And Not IsNull(rsTmp!住院包装) Then
                            .TextMatrix(i, col总量) = FormatEx(rsTmp!总量 / rsTmp!住院包装, 5)
                        End If
                        If NVL(rsTmp!医嘱期效, 0) = 1 Then
                            .TextMatrix(i, col总量单位) = NVL(rsTmp!住院单位)
                        End If
                    End If
                Else
                    If Not IsNull(rsTmp!总量) Then
                        .TextMatrix(i, col总量) = FormatEx(rsTmp!总量, 5)
                    End If
                    If NVL(rsTmp!医嘱期效, 0) = 1 Then
                        If NVL(rsTmp!类别) = "4" Then
                            .TextMatrix(i, col总量单位) = NVL(rsTmp!散装单位)
                        Else
                            .TextMatrix(i, col总量单位) = NVL(rsTmp!计算单位)
                        End If
                    End If
                End If
                
                .TextMatrix(i, col频次) = NVL(rsTmp!执行频次)
                .TextMatrix(i, col嘱托) = NVL(rsTmp!医生嘱托)
                .TextMatrix(i, col执行时间) = NVL(rsTmp!执行时间方案)
                .TextMatrix(i, col执行科室) = NVL(rsTmp!执行科室)
                .TextMatrix(i, col单独应用) = NVL(rsTmp!单独应用)
                .TextMatrix(i, col操作类型) = NVL(rsTmp!操作类型)
                .TextMatrix(i, col项目名称) = NVL(rsTmp!项目名称)
                .TextMatrix(i, col收费名称) = NVL(rsTmp!收费名称)
                
                '处理行隐藏及用法显示
                If InStr(",C,D,F,G,E,", NVL(rsTmp!类别, "*")) > 0 And Not IsNull(rsTmp!相关ID) Then
                    .RowHidden(i) = True
                    
                    '输血途径
                    If rsTmp!类别 = "E" And .TextMatrix(i - 1, col诊疗类别) = "K" _
                        And rsTmp!相关ID = Val(.TextMatrix(i - 1, colID)) Then
                        .TextMatrix(i - 1, col用法) = rsTmp!名称
                    End If
                ElseIf NVL(rsTmp!类别) = "7" Then
                    .RowHidden(i) = True
                ElseIf NVL(rsTmp!类别) = "E" And IsNull(rsTmp!相关ID) _
                    And Val(.TextMatrix(i - 1, col相关ID)) = rsTmp!ID _
                    And InStr(",5,6,", .TextMatrix(i - 1, col诊疗类别)) > 0 Then
                    '给药途径
                    .RowHidden(i) = True
                    '显示给药途径
                    For j = i - 1 To .FixedRows Step -1
                        If Val(.TextMatrix(j, col相关ID)) = rsTmp!ID Then
                            .TextMatrix(j, col用法) = rsTmp!名称 & rsTmp!医生嘱托
                        Else
                            Exit For
                        End If
                    Next
                ElseIf NVL(rsTmp!类别) = "E" And IsNull(rsTmp!相关ID) _
                    And Val(.TextMatrix(i - 1, col相关ID)) = rsTmp!ID _
                    And InStr(",7,E,C,", .TextMatrix(i - 1, col诊疗类别)) > 0 Then
                    '中药用法或检验采集方法
                    .TextMatrix(i, col用法) = rsTmp!名称
                    
                    '中药或检验的执行科室
                    For j = i - 1 To .FixedRows Step -1
                        If Val(.TextMatrix(j, col相关ID)) = rsTmp!ID Then
                            If InStr(",7,C,", .TextMatrix(j, col诊疗类别)) > 0 Then
                                .TextMatrix(i, col执行科室) = .TextMatrix(j, col执行科室)
                                Exit For
                            End If
                        Else
                            Exit For
                        End If
                    Next
                    
                    '中药付数
                    If .TextMatrix(i - 1, col诊疗类别) <> "C" Then
                        .TextMatrix(i, col总量单位) = "付"
                    End If
                End If
                
                '标记包含得有撤档或不服务的项目
                If Not IsNull(rsTmp!诊疗项目ID) Then
                    If Not (IsNull(rsTmp!撤档时间) Or Format(NVL(rsTmp!撤档时间), "yyyy-MM-dd") = "3000-01-01") Then
                        .RowData(i) = con项目撤档
                    ElseIf Not (NVL(rsTmp!服务对象, 0) = 3 Or NVL(rsTmp!服务对象, 0) = IIF(mstr挂号单 <> "", 1, 2)) Then
                        .RowData(i) = con项目服务
                    ElseIf Not IsNull(rsTmp!收费细目ID) Then
                        '对药品,同时要判断到收费项目目录
                        If Not (IsNull(rsTmp!收费撤档) Or Format(NVL(rsTmp!收费撤档), "yyyy-MM-dd") = "3000-01-01") Then
                            .RowData(i) = con收费撤档
                        ElseIf Not (NVL(rsTmp!收费服务, 0) = 3 Or NVL(rsTmp!收费服务, 0) = IIF(mstr挂号单 <> "", 1, 2)) Then
                            .RowData(i) = con收费服务
                        End If
                    ElseIf rsTmp!类别 <> rsTmp!诊疗类别 Then
                        .RowData(i) = con项目类别
                    End If
                End If
                
                If gblnStock Then
                    '判断非院外执行药品是否有库存
                    strDepartments = ""
                    If mlng病人科室id <> 0 And Val(rsTmp!执行性质 & "") <> 5 And InStr(",5,6,7,", rsTmp!类别 & "") > 0 And Val(rsTmp!收费细目ID & "") <> 0 Then
                        strDepartments = Get可用药房IDs(rsTmp!类别 & "", NVL(rsTmp!诊疗项目ID), Val(rsTmp!收费细目ID & ""), mlng病人科室id, mint来源)
                        '判断库存是否大于总量
                        If strDepartments <> "" Then
                            If GetStock(Val(rsTmp!收费细目ID & ""), , mint来源, strDepartments, CDbl(Val(.TextMatrix(i, col总量)))) = 0 Then
                                .Cell(flexcpData, i, col是否适用) = 1
                            End If
                        End If
                    End If
                End If

                rsTmp.MoveNext
            Next
        End If
        If mlng主页ID <> 0 Then
            .Cell(flexcpBackColor, .FixedRows, col选择, .Rows - 1, col期效) = COLEditBackColor      '浅绿
        End If
        
        .Row = .FixedRows: .Col = .FixedCols
        .AutoSize col内容
        .ColHidden(col期效) = mstr挂号单 <> ""
        .Redraw = flexRDDirect
    End With
    LoadAdvice = True
    Exit Function
errH:
    vsAdvice.Redraw = flexRDDirect
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub txtPati_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lvwList As ListItem, strTmp As String
    strTmp = lvwPati.SelectedItem.Text
    If KeyCode = 13 Then
        KeyCode = 0
        If IsNumeric(Trim(txtPati.Text)) And InStr(Trim(txtPati.Text), ".") <= 0 _
                And InStr(Trim(txtPati.Text), "-") <= 0 And InStr(Trim(txtPati.Text), "+") <= 0 Then
            Set lvwList = lvwPati.FindItem(Trim(txtPati.Text), 1)
        Else
            Set lvwList = lvwPati.FindItem(Trim(txtPati.Text))
        End If
        If lvwList Is Nothing Then
            MsgBox "没有找到该病人。", vbInformation, Me.Caption
            txtPati.Text = strTmp
            txtPati.SelStart = 0
            txtPati.SelLength = Len(txtPati.Text)
        Else
            lvwList.Selected = True
            Call lvwPati_KeyPress(13)
        End If
    End If
End Sub

Private Sub vsAdvice_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Dim lngW As Long
    
    If Col = col内容 Then
        vsAdvice.AutoSize Col
    ElseIf Row = -1 Then
        lngW = Me.TextWidth(vsAdvice.TextMatrix(vsAdvice.FixedRows - 1, Col) & "A")
        If vsAdvice.ColWidth(Col) < lngW Then
            vsAdvice.ColWidth(Col) = lngW
        ElseIf vsAdvice.ColWidth(Col) > vsAdvice.Width * 0.5 Then
            vsAdvice.ColWidth(Col) = vsAdvice.Width * 0.5
        End If
    End If
End Sub

Private Sub vsAdvice_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = col选择 Then Cancel = True
End Sub

Private Sub vsAdvice_DblClick()
    Call vsAdvice_KeyPress(32)
End Sub

Private Sub vsAdvice_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT
    
    With vsAdvice
        '擦除一并给药相关行列的边线及内容
        lngLeft = col期效: lngRight = col时间
        If Not Between(Col, lngLeft, lngRight) Then
            lngLeft = col频次: lngRight = col用法
            If Not Between(Col, lngLeft, lngRight) Then Exit Sub
        End If
        
        If Not RowIn一并给药(Row, lngBegin, lngEnd) Then Exit Sub
        
        vRect.Left = Left '擦除左边表格线
        vRect.Right = Right - 1 '保留右边表格线
        If Row = lngBegin Then
            vRect.Top = Bottom - 1 '首行保留文字内容
            vRect.Bottom = Bottom
        Else
            If Row = lngEnd Then
                vRect.Top = Top
                vRect.Bottom = Bottom - 1 '底行保留下边线
            Else
                vRect.Top = Top
                vRect.Bottom = Bottom
            End If
        End If
        If Between(Row, .Row, .RowSel) Then
            SetBkColor hDC, OS.SysColor2RGB(.BackColorSel)
        Else
            If Col = lngLeft And lngLeft = col期效 Then
                SetBkColor hDC, OS.SysColor2RGB(.Cell(flexcpBackColor, Row, lngLeft))
            Else
                SetBkColor hDC, OS.SysColor2RGB(.BackColor)
            End If
        End If
        ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        Done = True
    End With
End Sub

Private Sub vsAdvice_GotFocus()
    vsAdvice.BackColorSel = &HFFCC99
End Sub

Private Sub vsAdvice_KeyPress(KeyAscii As Integer)
    Dim i As Long
    
    With vsAdvice
        If KeyAscii = 32 Then
            If .Col <> col选择 Then
                KeyAscii = 0
                If .TextMatrix(.Row, col选择) = 0 Then
                    If CheckCanSelGroup(.Row, True) Then
                        If .Col = col期效 And mlng主页ID <> 0 And CanAlterType(.Row) Then
                            Call AlterGroupType(.Row)
                        End If
                        Call SelGroup(.Row, -1)
                    End If
                Else
                    If .Cell(flexcpFontBold, .Row, .Col) Then
                        Call AlterGroupType(.Row)
                    End If
                    Call SelGroup(.Row, 0)
                End If
            End If
        ElseIf KeyAscii = 13 Then
            KeyAscii = 0
            For i = .Row + 1 To .Rows - 1
                If Not .RowHidden(i) Then
                    .Row = i
                    Call .ShowCell(.Row, .Col)
                    Exit For
                End If
            Next
            If i > .Rows - 1 Then
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        End If
    End With
End Sub

Private Sub vsAdvice_LostFocus()
    vsAdvice.BackColorSel = &HC0C0C0
End Sub

Private Sub vsAdvice_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsAdvice
        If Col <> col选择 Then
            Cancel = True
        ElseIf Val(.TextMatrix(vsAdvice.Row, colID)) = 0 Then
            Cancel = True
        Else
            If .TextMatrix(Row, col选择) <> 0 Then
                Call SelGroup(Row, 0)
            Else
                If CheckCanSelGroup(Row, True) Then
                    Call SelGroup(Row, -1)
                End If
            End If
            '已经进行判断后选择，不需触发AfterEdit事件
            Cancel = True
        End If
    End With
End Sub

Private Function CanAlterType(ByVal lngRow As Long) As Boolean
'功能：判断指定的医嘱是否可以切换期效
'参数：lngRow=可见的医嘱行
'说明：允许切换期效的条件：
'   1.成长嘱：执行频率=0(可选频率),2(持续性)
'   2.成临嘱：执行频率=0(可选频率),1(一次性);药品必须指定了规格
    Dim rsMore As New ADODB.Recordset
    Dim strSQL As String, strType As String, i As Long
    Dim lngBegin As Long, lngEnd As Long
    
    With vsAdvice
        If Val(.TextMatrix(lngRow, colID)) = 0 Then
            CanAlterType = True: Exit Function
        ElseIf Val(.TextMatrix(lngRow, col诊疗项目ID)) = 0 Then
            '自由输入的可以切换
            CanAlterType = True: Exit Function
        ElseIf RowIn配方行(lngRow) Then
            '中药配方固定可以切换
            CanAlterType = True: Exit Function
        ElseIf RowIn检验行(lngRow) Then
            '检验以检验行为准判断
            lngRow = .FindRow(.TextMatrix(lngRow, colID), , col相关ID)
            If lngRow = -1 Then Exit Function
        End If
    
        strType = IIF(.TextMatrix(lngRow, col期效) = "长嘱", "临嘱", "长嘱")
        
        '以原始频率为准判断:因为可选择频率的可能已缺成一次性
        strSQL = "Select 执行频率 From 诊疗项目目录 Where ID=[1]"
        Set rsMore = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.TextMatrix(lngRow, col诊疗项目ID)))
        
        If strType = "长嘱" Then
            If InStr(",0,2,", NVL(rsMore!执行频率, 0)) = 0 Then Exit Function
        Else
            If InStr(",0,1,", NVL(rsMore!执行频率, 0)) = 0 Then Exit Function
            If InStr(",5,6,", .TextMatrix(lngRow, col诊疗类别)) > 0 Then
                Call GetRowScope(lngRow, lngBegin, lngEnd)
                For i = lngBegin To lngEnd
                    If InStr(",5,6,", .TextMatrix(i, col诊疗类别)) > 0 Then
                        If Val(.TextMatrix(i, col收费细目ID)) = 0 Then Exit Function
                    End If
                Next
            End If
        End If
    End With
    CanAlterType = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function RowIn检验行(ByVal lngRow As Long) As Boolean
'功能：判断指定行是否属于检验组合中的一行
'说明：不管行当前是否隐藏
    If lngRow = -1 Then Exit Function
    If vsAdvice.TextMatrix(lngRow, colID) = 0 Then Exit Function
    
    With vsAdvice
        If .TextMatrix(lngRow, col诊疗类别) = "E" And Val(.TextMatrix(lngRow, col相关ID)) = 0 Then
            '采集方法行
            If .TextMatrix(lngRow - 1, col诊疗类别) = "C" _
                And Val(.TextMatrix(lngRow - 1, col相关ID)) = .TextMatrix(lngRow, colID) Then
                RowIn检验行 = True: Exit Function
            End If
        ElseIf .TextMatrix(lngRow, col诊疗类别) = "C" And Val(.TextMatrix(lngRow, col相关ID)) <> 0 Then
            '检验项目行
            RowIn检验行 = True: Exit Function
        End If
    End With
End Function

Private Function RowIn配方行(ByVal lngRow As Long) As Boolean
'功能：判断指定行是否属于中药配方中的一行
'说明：不管行当前是否隐藏
    If lngRow = -1 Then Exit Function
    If vsAdvice.TextMatrix(lngRow, colID) = 0 Then Exit Function
    
    With vsAdvice
        If .TextMatrix(lngRow, col诊疗类别) = "E" Then
            If Val(.TextMatrix(lngRow, col相关ID)) = 0 Then
                '用法行
                If Val(.TextMatrix(lngRow - 1, col相关ID)) = .TextMatrix(lngRow, colID) _
                    And .TextMatrix(lngRow - 1, col诊疗类别) = "E" Then
                    RowIn配方行 = True: Exit Function
                End If
            Else
                '煎法行
                If .TextMatrix(lngRow - 1, col诊疗类别) = "7" _
                    And Val(.TextMatrix(lngRow - 1, col相关ID)) = Val(.TextMatrix(lngRow, col相关ID)) Then
                    RowIn配方行 = True: Exit Function
                End If
            End If
        ElseIf .TextMatrix(lngRow, col诊疗类别) = "7" And Val(.TextMatrix(lngRow, col相关ID)) <> 0 Then
            '中药行
            RowIn配方行 = True: Exit Function
        End If
    End With
End Function

Private Sub vsPati_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim i As Integer
    
    If vsPati.Redraw = flexRDNone Then Exit Sub
    If NewRow = -1 Then Exit Sub
    If NewRow = OldRow Then Exit Sub
    
    For i = 0 To cboTime.ListCount - 1
        If cboTime.ItemData(i) = vsPati.RowData(NewRow) Then
            cboTime.ListIndex = i: Exit For
        End If
    Next
    
    With vsPati
        If .RowData(NewRow) = -1 Then
            .Tag = "加载部分"
            For i = 4 To .Rows - 2
                .RowHidden(i) = False
            Next
            .RowHidden(.Rows - 1) = True
        End If
    End With
End Sub

Private Sub vsPati_GotFocus()
    vsPati.BackColorSel = &HFFCC99
End Sub

Private Sub vsPati_LostFocus()
    vsPati.BackColorSel = &HC0C0C0
End Sub

Private Function CheckCanSelRow(ByVal lngRow As Long) As String
'功能:验证指定行是否可以选择
    Dim lngCol As Long
    Dim strContent As String

    With vsAdvice
        If .TextMatrix(lngRow, col诊疗类别) = "D" And Trim(.Cell(flexcpData, lngRow, col内容) & "") <> "" Then
            If strContent <> "[]" Then
                strContent = Chr(34) & strContent & Chr(34)
            Else
                strContent = Chr(34) & .Cell(flexcpData, lngRow, col单量) & Chr(34)
            End If
        Else
            strContent = Chr(34) & .Cell(flexcpData, lngRow, col单量) & Chr(34)
        End If
        If Val(.RowData(lngRow)) < 0 Then
            Select Case Val(.RowData(lngRow))
            Case con项目撤档
                strContent = "诊疗项目【" & .TextMatrix(lngRow, col项目名称) & "】已撤档。"
            Case con项目服务
                strContent = "诊疗项目【" & .TextMatrix(lngRow, col项目名称) & "】服务对象发生变化。"
            Case con项目类别
                strContent = "诊疗项目【" & .TextMatrix(lngRow, col项目名称) & "】类别发生变化。"
            Case con收费撤档
                strContent = "收费项目【" & .TextMatrix(lngRow, col收费名称) & "】已撤档。"
            Case con收费服务
                strContent = "收费项目【" & .TextMatrix(lngRow, col收费名称) & "】服务对象发生变化。"
            End Select
            CheckCanSelRow = strContent: Exit Function
        End If

        If InStr("未知" & mstr性别, .TextMatrix(lngRow, col性别)) = 0 Then
            CheckCanSelRow = strContent & "(不适用于当前病人性别)": Exit Function
        End If
        
        If .TextMatrix(lngRow, col是否适用) = "1" Then
            CheckCanSelRow = strContent & "(不适用于当前科室)": Exit Function
        End If

        If mbln麻醉类权限 And .TextMatrix(lngRow, col毒理分类) = "麻醉药" Then
            CheckCanSelRow = strContent & "(无麻醉类药品权限)": Exit Function

        End If

        If mbln毒性类权限 And .TextMatrix(lngRow, col毒理分类) = "毒性药" Then
            CheckCanSelRow = strContent & "(无毒性药品权限)": Exit Function

        End If

        If mbln精神类权限 And (.TextMatrix(lngRow, col毒理分类) = "精神I类") Then
            CheckCanSelRow = strContent & "(无精神类药品权限)": Exit Function

        End If

        If mbln贵重类权限 And (.TextMatrix(lngRow, col价值分类) = "贵重" Or .TextMatrix(lngRow, col价值分类) = "昂贵") Then
            CheckCanSelRow = strContent & "(无贵重类药品权限)": Exit Function

        End If
        
        '输血医嘱检查，必须中级及以上专业技术职务的医师才允许下达
        If .TextMatrix(lngRow, col诊疗类别) = "K" And gbln输血申请中级以上 Then
            If UserInfo.专业技术职务 <> "主治医师" And UserInfo.专业技术职务 <> "主任医师" And UserInfo.专业技术职务 <> "副主任医师" Then
                CheckCanSelRow = Trim(.Cell(flexcpText, lngRow, col内容) & "") & "(无中级及以上专业技术职务)": Exit Function
            End If
        End If

    End With

End Function

Private Sub GetRowScope(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long, Optional lngRow相关 As Long)
'功能:获取一组医嘱的起止位置，同时获取父医嘱行号
'参数:
'   lngRow 当前行
'返回:
'   lngBegin 起始行
'   lngEnd 终止行
'   lngRow相关 父医嘱行

    Dim i As Long, lng相关 As Long

    With vsAdvice
        If .TextMatrix(lngRow, col诊疗类别) = "" Then  '自由录入
            lngRow相关 = lngRow: lngBegin = lngRow: lngEnd = lngRow
            Exit Sub
        End If
        '获取相关序号
        If Val(.TextMatrix(lngRow, col相关ID)) <> 0 Then
            lng相关 = Val(.TextMatrix(lngRow, col相关ID))
            lngRow相关 = .FindRow(lng相关, , colID, , True)
            If lngRow相关 = -1 Then
                lngRow相关 = lngRow
            End If
        Else
            lng相关 = Val(.TextMatrix(lngRow, colID)): lngRow相关 = lngRow
        End If

        lngBegin = lngRow相关: lngEnd = lngRow相关

        For i = lngRow相关 - 1 To .FixedRows Step -1
            If Not (Val(.TextMatrix(i, colID)) = 0 And i >= .FixedRows) Then '跳过（一并给药中）空行
                If Val(.TextMatrix(i, col相关ID)) = lng相关 Then
                    lngBegin = i
                Else
                    Exit For
                End If
            End If
        Next

        For i = lngRow相关 + 1 To .Rows - 1
            If Not (Val(.TextMatrix(i, colID)) = 0 And i >= .FixedRows) Then '跳过（一并给药中）空行
                If Val(.TextMatrix(i, col相关ID)) = lng相关 Then
                    lngEnd = i
                Else
                    Exit For
                End If
            End If
        Next
    End With
End Sub

Private Function CheckCanSelGroup(ByVal lngRow As Long, Optional ByVal blnAsk As Boolean = True) As Boolean
'功能：判断本组医嘱是否可以选择
'参数：
'   lngRow 当前行
'   blnAsk 是否进行提示或询问（全选时，不进行询问)
    Dim i As Long, strResult As String
    Dim lngBegin As Long, lngEnd As Long, lngRow相关 As Long
    Dim bln配方 As Boolean, blnCanSel As Boolean
    Dim strMsg As String
    Dim blnMedicineAdvice As Boolean

    With vsAdvice
        '获取本组医嘱信息
        Call GetRowScope(lngRow, lngBegin, lngEnd, lngRow相关)
        
        If Not mbln医技后续 And mlng前提ID <> 0 Then
            If .TextMatrix(lngRow, col期效) = "长嘱" Then
                If blnAsk Then
                    MsgBox "系统不允许对医技医嘱进行后续处理，不允许复制长嘱！", vbInformation, gstrSysName
                End If
                Exit Function
            End If
        End If
        
         '检查是否有医嘱的诊疗项目未勾选“可以单独应用”，未勾选的不允许复制。
        If lngBegin = lngEnd Then
            If Val(.TextMatrix(lngRow, col单独应用)) = 0 And Val(.TextMatrix(lngRow, col诊疗项目ID)) <> 0 Then
                If blnAsk Then
                    MsgBox "医嘱“" & .TextMatrix(lngRow, col内容) & "”对应的诊疗项目不能单独应用，不可以被复制。如有疑问，请联系管理员！", vbInformation, gstrSysName
                End If
                Exit Function
            End If
        Else
            For i = lngBegin To lngEnd
                If InStr(",5,6,7,", .TextMatrix(i, col诊疗类别)) > 0 Then
                    blnMedicineAdvice = True
                End If
            Next
            If Not blnMedicineAdvice Then
                For i = lngBegin To lngEnd
                    If Not (.TextMatrix(i, col诊疗类别) = "G" Or (.TextMatrix(i, col诊疗类别) = "E" And InStr(",2,3,4,6,7,8,", .TextMatrix(i, col操作类型)) > 0)) Then
                        If Val(.TextMatrix(i, col单独应用)) = 0 And Val(.TextMatrix(i, col诊疗项目ID)) <> 0 Then
                            If blnAsk Then
                                MsgBox "医嘱“" & .TextMatrix(i, col内容) & "”对应的诊疗项目不能单独应用，不可以被复制。如有疑问，请联系管理员！", vbInformation, gstrSysName
                            End If
                            Exit Function
                        End If
                    End If
                Next
            End If
        End If

        
        strMsg = ""
        '在启用参数 指定药房时限制库存 的情况下，不允许复制库存不足的医嘱
        If gblnStock Then
            For i = lngBegin To lngEnd
                If Val(.Cell(flexcpData, i, col是否适用)) = 1 Then
                    If Val(.TextMatrix(lngBegin, col诊疗类别)) = 7 Then
                        strMsg = strMsg & "," & .TextMatrix(i, col内容)
                    Else
                        If blnAsk Then
                            MsgBox "该药品库存不足,系统限制了不允许下达库存不足的药品，不能被选择！", vbInformation, gstrSysName
                        End If
                        Exit Function
                    End If
                End If
            Next
        End If
        If strMsg <> "" Then
            MsgBox "该配方中存在库存不足的药品(" & Mid(strMsg, 2) & ")。", vbInformation, gstrSysName
        End If
        
        strMsg = CheckCanSelRow(lngRow相关)
        If strMsg <> "" Then '父医嘱或单条医嘱检查
            If blnAsk Then
                MsgBox "该医嘱中:" & vbNewLine & strMsg & vbNewLine & "无效,不能被选择", vbInformation, gstrSysName
            End If
            Exit Function
        Else
            If lngBegin <> lngEnd Then
                If .TextMatrix(lngRow相关, col诊疗类别) = "E" Then
                    If lngRow相关 - 2 >= lngBegin Then
                        If .TextMatrix(lngRow相关 - 2, col诊疗类别) = "7" And .TextMatrix(lngRow相关 - 1, col诊疗类别) = "E" Then '中药配方的剪法检查
                            strMsg = CheckCanSelRow(lngRow相关 - 1)
                            If strMsg <> "" Then
                                If blnAsk Then
                                    MsgBox "该中药配方中煎法:" & vbNewLine & strMsg & vbNewLine & "无效,不能被选择", vbInformation, gstrSysName
                                End If
                                Exit Function
                            Else
                                bln配方 = True
                            End If
                        End If
                    End If
                End If
            End If
        End If
        strMsg = ""
        '子医嘱全部检查
        If lngBegin <> lngEnd Then
            For i = lngBegin To lngEnd
                If .TextMatrix(lngRow相关, col诊疗类别) = "F" Then blnCanSel = True '手术医嘱父医嘱可用就可选
                If Not (i = lngRow相关 Or bln配方 And i = lngRow相关 - 1) Then
                    strResult = CheckCanSelRow(i)
                    If strResult <> "" Then
                        strMsg = IIF(strMsg = "", "", strMsg & "、" & vbNewLine) & strResult
                    Else
                        If bln配方 Then  '中药配方含一味中药可用就可选（煎法以及用法前面已经判断），其余类型只要一个子医嘱可用就可选
                            If .TextMatrix(i, col诊疗类别) = "7" Then
                                blnCanSel = True
                            End If
                        Else
                            blnCanSel = True
                        End If
                    End If
                End If
            Next
        Else
            blnCanSel = True '单条医嘱检查在父医嘱时已经检查
        End If
        
        If Not blnCanSel Then
        '中药配方未提取的药品信息
            If bln配方 Then
                strMsg = "该中药配方中所有中药已经被停用或没有可用规格,不能被选择"
            Else
                If strMsg = "" Then strMsg = "该医嘱中不存在有效项目,不能被选择"
            End If
        End If

        If blnCanSel Then
            If strMsg <> "" Then
                If blnAsk Then
                    If MsgBox(IIF(InStr(1, strMsg, "、") > 0, "该医嘱中:" & vbNewLine & strMsg & vbNewLine & "无效,这些项目", "该医嘱中:" & vbNewLine & strMsg & vbNewLine & "无效,该项目") & "不会被选择,是否选择该医嘱？", _
                        vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
                        CheckCanSelGroup = True
                    End If
                End If
            Else
                CheckCanSelGroup = True
            End If
        Else
            If blnAsk Then
                MsgBox strMsg, vbInformation, gstrSysName
            End If
        End If
    End With
End Function

Private Sub SelGroup(ByVal lngRow As Long, ByVal int选择 As Integer, Optional ByRef lngEnd As Long)
'功能:根据情况选择该组医嘱
'参数：
'   lngRow 当前行
'   lngEnd 本组医嘱最后一行
'   int选择 选择结果 -1,检查选择（可选的选择，不可选不选择),0不选择,1，全选不检查
    Dim lngBegin As Long
    Dim i As Long

    With vsAdvice
        
        '获取本组医嘱信息
        Call GetRowScope(lngRow, lngBegin, lngEnd)
        '选择或取消选择
        If int选择 = -1 Then 'checkCanSelGroup(i,true)=true后调用
            For i = lngBegin To lngEnd
                If CheckCanSelRow(i) = "" Then
                    .TextMatrix(i, col选择) = int选择
                Else
                    .TextMatrix(i, col选择) = 0
                End If
            Next
        Else 'checkCanSelGroup(i,false)=true 后调用或取消选择后使用
            int选择 = int选择 * -1
            For i = lngBegin To lngEnd
                .TextMatrix(i, col选择) = int选择
            Next
        End If
    End With
End Sub

Private Sub AlterGroupType(ByVal lngRow As Long)
'功能：改变指定行所在医嘱组的医嘱期效
    Dim lngBegin As Long, lngEnd As Long, i As Long
    
    With vsAdvice
        .Redraw = False
        '获取本组医嘱信息
        Call GetRowScope(lngRow, lngBegin, lngEnd)
        '改变本组所有医嘱期效
        For i = lngBegin To lngEnd
            If .TextMatrix(i, .Col) <> "" Then
                .TextMatrix(i, .Col) = IIF(.TextMatrix(i, .Col) = "长嘱", "临嘱", "长嘱")
            End If
            .Cell(flexcpFontBold, i, .Col) = .TextMatrix(i, .Col) <> .Cell(flexcpData, i, .Col)
        Next
        .Redraw = True
    End With
End Sub

Private Sub InitAdviceTable()
'功能：初始化表格内容
    Dim strHead As String, i As Integer
    Dim arrHead As Variant, arrCol As Variant
    
    strHead = ",300,4;期效,480,1;开始时间,1530,1;内容,2910,1;总量,630,7;单位,465,1;单量,630,7;单位,465,1;频次,1140,1;用法,1140,1;医生嘱托,1650,1;执行时间,1095,1;执行科室,1110,1;" & _
        "ID;相关ID;诊疗类别;诊疗项目ID;收费细目ID;是否适合;毒理分类;价值分类;性别;单独应用;操作类型;项目名称;收费名称"
    arrHead = Split(strHead, ";")
    With vsAdvice
        .Clear
        .FixedRows = 1: .FixedCols = 0
        .Rows = 2: .Cols = .FixedCols + UBound(arrHead) + 1
        For i = 0 To UBound(arrHead)
            .FixedAlignment(.FixedCols + i) = 4
            arrCol = Split(arrHead(i), ",")
            .TextMatrix(0, .FixedCols + i) = arrCol(0)
            If UBound(arrCol) > 0 Then
                .ColWidth(.FixedCols + i) = Val(arrCol(1))
                .ColAlignment(.FixedCols + i) = Val(arrCol(2))
                .ColHidden(.FixedCols + i) = False
            Else
                .ColHidden(.FixedCols + i) = True
            End If
        Next
    End With
End Sub

