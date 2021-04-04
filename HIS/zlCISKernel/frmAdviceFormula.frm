VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmAdviceFormula 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   4395
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7035
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   7035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox txtJL 
      Height          =   300
      Left            =   2955
      MaxLength       =   20
      TabIndex        =   24
      Top             =   3690
      Width           =   900
   End
   Begin VB.ComboBox cboData 
      Height          =   300
      Left            =   480
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   3720
      Visible         =   0   'False
      Width           =   2000
   End
   Begin VB.OptionButton optMode 
      Caption         =   "散装(&0)"
      Enabled         =   0   'False
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   20
      Top             =   97
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.OptionButton optMode 
      Caption         =   "饮片(&1)"
      Enabled         =   0   'False
      Height          =   180
      Index           =   1
      Left            =   1155
      TabIndex        =   19
      Top             =   97
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.OptionButton optMode 
      Caption         =   "免煎剂(&2)"
      Enabled         =   0   'False
      Height          =   180
      Index           =   2
      Left            =   2310
      TabIndex        =   18
      Top             =   97
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.Frame fraBorder 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   45
      Index           =   0
      Left            =   3210
      MousePointer    =   7  'Size N S
      TabIndex        =   17
      Top             =   3975
      Width           =   615
   End
   Begin VB.Frame fraBorder 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   45
      Index           =   2
      Left            =   3210
      TabIndex        =   16
      Top             =   4245
      Width           =   615
   End
   Begin VB.Frame fraBorder 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   330
      Index           =   3
      Left            =   3120
      TabIndex        =   15
      Top             =   3960
      Width           =   45
   End
   Begin VB.Frame fraBorder 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   330
      Index           =   1
      Left            =   3870
      MousePointer    =   9  'Size W E
      TabIndex        =   14
      Top             =   3975
      Width           =   45
   End
   Begin VB.Frame fraBorder 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   45
      Index           =   4
      Left            =   4185
      MousePointer    =   7  'Size N S
      TabIndex        =   13
      Top             =   4110
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Frame fra中药 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   3000
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   3975
      Begin VB.ComboBox cbo药房 
         Height          =   300
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   0
         Width           =   1920
      End
      Begin VB.TextBox txt付数 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1080
         MaxLength       =   3
         TabIndex        =   9
         Text            =   "1"
         Top             =   0
         Width           =   450
      End
      Begin VB.Label lbl药房 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "药房"
         Height          =   240
         Left            =   1680
         TabIndex        =   12
         Top             =   60
         Width           =   405
      End
      Begin VB.Label lbl付数 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "付数"
         Height          =   240
         Left            =   720
         TabIndex        =   11
         Top             =   60
         Width           =   405
      End
   End
   Begin VB.CommandButton cmdOK 
      Height          =   315
      Left            =   5895
      Picture         =   "frmAdviceFormula.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "确认(F2)"
      Top             =   3720
      Width           =   450
   End
   Begin VB.CommandButton cmdCancel 
      Height          =   315
      Left            =   6450
      Picture         =   "frmAdviceFormula.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "取消(Esc)"
      Top             =   3720
      Width           =   450
   End
   Begin VB.CommandButton cmdInsert 
      Height          =   315
      Left            =   5400
      Picture         =   "frmAdviceFormula.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "插入(&A)"
      Top             =   3720
      Visible         =   0   'False
      Width           =   450
   End
   Begin VSFlex8Ctl.VSFlexGrid vsExt 
      Height          =   1845
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   6900
      _cx             =   12171
      _cy             =   3254
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
      BackColorSel    =   4210752
      ForeColorSel    =   16777215
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   3
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   7
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmAdviceFormula.frx":7366
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
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
      Begin VB.CommandButton cmd 
         Caption         =   "…"
         Height          =   240
         Left            =   4920
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "选择项目(*)"
         Top             =   720
         Visible         =   0   'False
         Width           =   270
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vs中药规格 
      Height          =   1335
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2280
      Visible         =   0   'False
      Width           =   6975
      _cx             =   12303
      _cy             =   2355
      Appearance      =   0
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
      BackColorSel    =   4210752
      ForeColorSel    =   16777215
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   4
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmAdviceFormula.frx":7472
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
      ScrollTips      =   0   'False
      MergeCells      =   1
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
      OwnerDraw       =   0
      Editable        =   2
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
      Begin VB.CommandButton cmd形态 
         Caption         =   "散装(&D)"
         Height          =   255
         Left            =   2760
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Label lblJL 
      Caption         =   "煎量"
      Height          =   240
      Left            =   2565
      TabIndex        =   25
      Top             =   3705
      Width           =   405
   End
   Begin VB.Line lin 
      Index           =   0
      X1              =   3840
      X2              =   4515
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line lin 
      Index           =   1
      X1              =   3840
      X2              =   4515
      Y1              =   3750
      Y2              =   3750
   End
   Begin VB.Line lin 
      Index           =   2
      X1              =   3840
      X2              =   4515
      Y1              =   3780
      Y2              =   3780
   End
   Begin VB.Line lin 
      Index           =   3
      X1              =   3840
      X2              =   4515
      Y1              =   3810
      Y2              =   3810
   End
   Begin VB.Line lin 
      Index           =   4
      X1              =   3840
      X2              =   4515
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line lin 
      Index           =   5
      X1              =   3840
      X2              =   4515
      Y1              =   3870
      Y2              =   3870
   End
   Begin VB.Line lin 
      Index           =   6
      X1              =   3840
      X2              =   4515
      Y1              =   3900
      Y2              =   3900
   End
   Begin VB.Line lin 
      Index           =   7
      X1              =   3840
      X2              =   4515
      Y1              =   3930
      Y2              =   3930
   End
   Begin VB.Label lblData 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "煎法"
      Height          =   180
      Left            =   105
      TabIndex        =   7
      Top             =   3780
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblNumZY 
      Caption         =   "共___味"
      Height          =   255
      Left            =   4560
      TabIndex        =   4
      Top             =   3840
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lblZYStock 
      Caption         =   "中药库存显示"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   4080
      Visible         =   0   'False
      Width           =   3975
   End
End
Attribute VB_Name = "frmAdviceFormula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'=============================================================================================================
'入口参数：
Private mclsInsure As Object '该窗体为公共，可能其他工程中不使用医保，也不引用
Private mlngHwnd As Long '用于定位的控件句柄
Private mint期效 As Integer
Private mstr性别 As String
Private mint调用类型 As Integer  '1-门诊,2-住院
Private mint服务对象 As Integer '1-门诊,2-住院,3-门诊和住院
Private mbytUseType As Byte      '0=医嘱下达,1-路径项目的医嘱生成,2-添加路径外项目

'入:主诊疗项目ID,中药配方时为配方ID或单味中药ID
Private mlng项目ID As Long


'入/出:附加定义数据,新增时一般为空
'      中药="中药ID1,单量1,脚注1;中药ID2,单量2,脚注2;...|煎法ID|中药形态|付数|药房ID|煎量"
Private mstrExtData As String
Private mstr配方明细 As String

'入:部份情况下需要,如检查申请取附项内容
Private mlng病人ID As Long
Private mvar就诊ID As Variant '主页ID或挂号单号
Private mint婴儿 As Integer
Private mint险类 As Integer '医保病人的险类
Private mlng病人科室id As Long '用于确定中药配方的缺省药房
Private mlngPreRow中药房 As Long '上或下一行中药配方的药房
Private mlng药品ID As Long       '选择器选中的中药规格ID
Private mstr药品价格等级 As String '病人的药品价格等级
Private mlng病人性质 As Long

Private mint场合 As Integer  '0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS),3-成套方案(路径项目设置)调用
Private mbln医保 As Boolean '是否医保或公费病人

'出：由医保接口GetItemInfo所返回的摘要，主要处理配方的
Private mstr摘要 As String

'出口参数：
Private mblnOK As Boolean '出

Private mlng中药房 As Long
Private mstr可用中药房 As String

Private mblnFirst As Boolean
Private mblnReturn As Boolean '是否了回车确认
Private mcol规格数量 As Collection  '以药名ID为键存储药品的规格数量：规格1,数量;规格2,数量|未分配数量
Private mbytSize As Byte '字体大小 0-小字体（9号），1-大字体（12号）
Private mint每行味数 As Integer '中药配方每行中药的味数
Private Enum E中药规格
    col规格 = 0
    col产地 = 1
    col剂量单位 = 2
    col数量 = 3
    col单价 = 4
    col药品ID = 5
End Enum
Private mblnChangeSel As Boolean
Private mstrPrivs As String             '权限
Private mfrmParent As Object
Private mblnSelf As Boolean

Public Function ShowMe(ByVal frmParent As Object, ByVal objclsInsure As Object, ByVal lngHwnd As Long, ByRef t_Pati As TYPE_PatiInfoEx, ByVal int场合 As Integer, _
             ByVal bytUseType As Byte, ByVal int期效 As Integer, ByVal int服务对象 As Integer, Optional ByVal int调用类型 As Integer, _
             Optional ByVal lng项目id As Long, Optional ByRef strExtData As String, _
             Optional ByRef str摘要 As String, Optional ByVal lng药品ID As Long, Optional ByVal lngPreRow中药房 As Long, Optional ByVal str药品价格等级 As String) As Boolean
'参数:
'     frmParent         父窗体
'     objclsInsure      该窗体为公共，可能其他工程中不使用医保，也不引用,因此要传入医保部件
'     lngHwnd           用于定位的控件句柄,即调用该窗体的控件
'     t_Pati            病人信息
'     int场合           0-医生站调用,1-护士站调用,2-医技站调用(PACS/LIS),3-成套方案(路径项目设置)调用
'     bytUseType        0=医嘱下达,1-路径项目的医嘱生成,2-添加路径外项目,3-路径项目批量调整
'     int期效           将要输入的医嘱期效 0-长嘱，1-临嘱
'     int服务对象       该医嘱要服务的病人性质 1-门诊（包括门诊病人，体检病人，外来病人等) 2-住院（只有住院病人）
'     int调用类型       调用该窗体的工作站类型 1-门诊医生工作站 2-住院医护工作站(现在只针对中药配方)
'     lng项目id         主诊疗项目ID , 中药配方时为配方ID或单味中药ID
'     lng药品ID         选择器选中的中药规格ID
'     lngPreRow中药房   默认中药房，上或下一行中药配方的药房
'返回：
'     strExtData        附加定义数据 , 新增时一般为空
'                       中药 = "中药ID1,单量1,脚注1;中药ID2,单量2,脚注2;...|煎法ID|中药形态|付数|药房ID"
'     str摘要           由医保接口GetItemInfo所返回的摘要，主要处理配方的。

    Set mfrmParent = frmParent
    Set mclsInsure = objclsInsure
    mlngHwnd = lngHwnd
    With t_Pati
        mbln医保 = .bln医保
        mint险类 = .int险类
        mint婴儿 = .int婴儿
        mlng病人ID = .lng病人ID
        mlng病人科室id = .lng病人科室ID
        mvar就诊ID = IIF(.str挂号单 = "", .lng主页ID, .str挂号单)
        mstr性别 = .str性别
    End With
    mint场合 = int场合
    mbytUseType = bytUseType
    mint期效 = int期效
    mint服务对象 = int服务对象
    If mint场合 <> 3 Then
        mint调用类型 = int调用类型
    Else
        mint调用类型 = IIF(int服务对象 = 1, 1, 2)
    End If
    mlng项目ID = lng项目id
    mstrExtData = strExtData
    mstr摘要 = str摘要
    mlng药品ID = lng药品ID
    mlngPreRow中药房 = lngPreRow中药房
    mstr药品价格等级 = str药品价格等级
    mblnOK = False
    mlng病人性质 = 0 '内部进行赋值
    
    On Error Resume Next
    Me.Show 1, frmParent
    err.Clear: On Error GoTo 0
    
    strExtData = mstrExtData
    str摘要 = mstr摘要
    
    
    ShowMe = mblnOK
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmd形态_Click()
'功能：单味药，非散装分配不完时，换成散装规格
    Dim lng药名ID As Long, dbl数量 As Double, lng药品ID As Long
    Dim strKey As String
    
    strKey = vsExt.Cell(flexcpData, vsExt.Row, (vsExt.Col \ 4) * 4 + 2)
    If strKey <> "" Then lng药名ID = Val(Split(strKey, "_")(0))
    
    dbl数量 = Val(vsExt.TextMatrix(vsExt.Row, (vsExt.Col \ 4) * 4 + 1))
    lng药品ID = Val(cmd形态.Tag)    '缺省规格
        
    Call mcol规格数量.Remove("_" & strKey)
    mcol规格数量.Add lng药品ID & "," & dbl数量, "_" & strKey
    
    Call Show中药规格(lng药名ID, dbl数量, 0)
    mblnChangeSel = True
    vsExt.SetFocus
    mblnChangeSel = False
End Sub

Private Sub Form_Resize()
    Dim lngAppend As Long
    Dim lngMinRows As Long
    Dim lngRows As Long, i As Long
    Dim lngHeight As Long, lngTotalHeight As Long
    
    On Error Resume Next
    
    fraBorder(0).Left = 0
    fraBorder(0).Top = 0
    fraBorder(0).Width = Me.ScaleWidth
    fraBorder(1).Top = fraBorder(0).Top + fraBorder(0).Height
    fraBorder(1).Left = Me.ScaleWidth - fraBorder(1).Width
    fraBorder(1).Height = Me.ScaleHeight - fraBorder(0).Height * 2
    fraBorder(2).Left = 0
    fraBorder(2).Top = Me.ScaleHeight - fraBorder(2).Height
    fraBorder(2).Width = Me.ScaleWidth
    fraBorder(3).Top = fraBorder(0).Top + fraBorder(0).Height
    fraBorder(3).Left = 0
    fraBorder(3).Height = Me.ScaleHeight - fraBorder(0).Height * 2
    
    vsExt.Left = fraBorder(3).Width
    vsExt.Top = fraBorder(0).Top + fraBorder(0).Height + fra中药.Height
    vsExt.Width = Me.ScaleWidth - fraBorder(3).Width * 2

    fra中药.Left = vsExt.Left
    fra中药.Top = fraBorder(0).Top + fraBorder(0).Height
    fra中药.Width = vsExt.Width
    
    If Me.Visible = False Then
        For i = 0 To optMode.Count - 1
            Set optMode(i).Container = fra中药
            optMode(i).Top = lbl付数.Top
        Next
        optMode(0).Left = 60
        optMode(1).Left = optMode(0).Left + optMode(0).Width
        optMode(2).Left = optMode(1).Left + optMode(1).Width
    
        lbl付数.Left = optMode(2).Left + optMode(2).Width + 360
        txt付数.Left = lbl付数.Left + lbl付数.Width
        lbl药房.Left = txt付数.Left + txt付数.Width + 360
        cbo药房.Left = lbl药房.Left + lbl药房.Width
    End If
    
    vsExt.Height = Me.ScaleHeight - fraBorder(2).Height * 2 - (cboData.Height + 150) - fra中药.Height - vs中药规格.Height - IIF(lblZYStock.Visible, lblZYStock.Height + 60, 0)
    lngMinRows = 7
    With vsExt
        For i = .FixedRows To .Rows - 1
            If Replace(.Cell(flexcpText, i, 0, i, .Cols - 1), Chr(9), "") <> "" Then
                lngMinRows = i + .FixedRows
            Else
                Exit For
            End If
        Next
        lngRows = Int((vsExt.Height - vsExt.RowHeight(0) - 15) / (vsExt.RowHeight(1) + 15))
        If lngRows < lngMinRows Then lngRows = lngMinRows
        .Rows = lngRows
    End With
    Call SetSplitLine
    
    With vs中药规格
        .Top = vsExt.Top + vsExt.Height + 30
        .Left = vsExt.Left
        .Width = vsExt.Width
        .ColWidth(col产地) = .Width - .ColWidth(col规格) - .ColWidth(col剂量单位) - .ColWidth(col数量) - .ColWidth(col单价)
    End With

    
    
    '中药显示库存
    lblZYStock.Top = vs中药规格.Top + vs中药规格.Height + 60
    lblZYStock.Left = vs中药规格.Left
    lblZYStock.Width = vs中药规格.Width
    cboData.Top = vs中药规格.Top + vs中药规格.Height + 60 + IIF(mint场合 <> 3, lblZYStock.Height, 0) + 60
    lblData.Top = cboData.Top + (cboData.Height - lblData.Height) / 2
    cmdOK.Top = cboData.Top + (cboData.Height - cmdOK.Height) / 2
    cmdCancel.Top = cmdOK.Top
        

    lblData.Left = 200
    cboData.Left = lblData.Left + lblData.Width + fraBorder(3).Width
    cmdCancel.Left = Me.ScaleWidth - cmdCancel.Width - cmdCancel.Height
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - fraBorder(1).Width * 3
        
    cboData.Width = cbo药房.Width
    lblJL.Width = lblData.Width
    lblJL.Top = lblData.Top
    lblJL.Left = cboData.Left + cboData.Width + 200
    
    txtJL.Top = cboData.Top
    txtJL.Left = lblJL.Left + lblJL.Width
    
    cmdInsert.Top = cmdOK.Top
    cmdInsert.Left = cmdOK.Left - cmdInsert.Width - 100
    lblNumZY.Top = cmdOK.Top + 45
    lblNumZY.Left = cmdInsert.Left - lblNumZY.Width - 100
    
    txtJL.Width = lblNumZY.Left - txtJL.Left - 400
    
    Me.Refresh
End Sub

Private Sub Form_Activate()
    If mblnFirst And vsExt.TabStop And vsExt.Enabled And vsExt.Visible And Not Me.ActiveControl Is vsExt Then
        mblnFirst = False: vsExt.SetFocus '？不清楚为什么自动定位到rtfAppend上面去了。
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    
    If KeyCode = vbKeyEscape Then
        Call cmdCancel_Click
    ElseIf KeyCode = vbKeyF2 Then
        If cmdOK.Enabled And cmdOK.Visible Then Call cmdOK_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr(",;|'", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0 '不允许输入分隔符及单引号
    End If
End Sub

Private Sub Form_Load()
    Dim blnMulti As Boolean, vRect As RECT
    Dim str方法 As String, i As Long, lngBaseHeight As Long
    
    Me.Height = 2325
    
    '边框设置
    For i = 0 To fraBorder.UBound
        fraBorder(i).BackColor = vbButtonFace
    Next
    Set lin(0).Container = fraBorder(0): Set lin(1).Container = fraBorder(0)
    Set lin(2).Container = fraBorder(1): Set lin(3).Container = fraBorder(1)
    Set lin(4).Container = fraBorder(2): Set lin(5).Container = fraBorder(2)
    Set lin(6).Container = fraBorder(3): Set lin(7).Container = fraBorder(3)
    lin(0).X1 = 0: lin(0).Y1 = 0: lin(0).X2 = Screen.Width: lin(0).Y2 = lin(0).Y1: lin(0).BorderColor = &H8000000F
    lin(1).X1 = 0: lin(1).Y1 = Screen.TwipsPerPixelY: lin(1).X2 = Screen.Width: lin(1).Y2 = lin(1).Y1: lin(1).BorderColor = &H8000000E
    lin(2).X1 = fraBorder(1).Width - Screen.TwipsPerPixelX: lin(2).Y1 = 0: lin(2).X2 = lin(2).X1: lin(2).Y2 = Screen.Height: lin(2).BorderColor = &H80000011
    lin(3).X1 = fraBorder(1).Width - Screen.TwipsPerPixelX * 2: lin(3).Y1 = 0: lin(3).X2 = lin(3).X1: lin(3).Y2 = Screen.Height: lin(3).BorderColor = &H80000010
    lin(4).X1 = 0: lin(4).Y1 = fraBorder(2).Height - Screen.TwipsPerPixelY: lin(4).X2 = Screen.Width: lin(4).Y2 = lin(4).Y1: lin(4).BorderColor = &H80000011
    lin(5).X1 = 0: lin(5).Y1 = fraBorder(2).Height - Screen.TwipsPerPixelY * 2: lin(5).X2 = Screen.Width: lin(5).Y2 = lin(5).Y1: lin(5).BorderColor = &H80000010
    lin(6).X1 = 0: lin(6).Y1 = 0: lin(6).X2 = lin(6).X1: lin(6).Y2 = Screen.Height: lin(6).BorderColor = &H8000000F
    lin(7).X1 = Screen.TwipsPerPixelX: lin(7).Y1 = 0: lin(7).X2 = lin(7).X1: lin(7).Y2 = Screen.Height: lin(7).BorderColor = &H8000000E
    
    If mint服务对象 = 0 Then mint服务对象 = IIF(mint场合 = 3, 3, 2) '缺省为住院,成套缺省为住院和门诊
    mblnOK = False
    mblnFirst = True
    
    If mint场合 = 0 Then
        If mint调用类型 = 1 Then
            mbytSize = zlDatabase.GetPara("字体", glngSys, pm门诊医生站, "0")
        Else
            mbytSize = zlDatabase.GetPara("字体", glngSys, pm住院医生站, "0")
        End If
    ElseIf mint场合 = 1 Then
        mbytSize = zlDatabase.GetPara("字体", glngSys, pm住院护士站, "0")
    ElseIf mint场合 = 2 Then
        mbytSize = zlDatabase.GetPara("字体", glngSys, pm医技工作站, "0")
    End If

    
    '初始化表格样式
    mint每行味数 = IIF(Val(zlDatabase.GetPara(213, glngSys)) = 4, 4, 3)
    mstr摘要 = ""
    Set mcol规格数量 = New Collection
    mlng中药房 = Val(zlDatabase.GetPara(IIF(mint调用类型 = 2, "住院", "门诊") & "缺省中药房", glngSys, IIF(mint调用类型 = 2, pm住院医嘱下达, pm门诊医嘱下达), , , , , mlng病人科室id))
    mstr可用中药房 = zlDatabase.GetPara(IIF(mint调用类型 = 2, "住院", "门诊") & "可用中药房", glngSys, IIF(mint调用类型 = 2, pm住院医嘱下达, pm门诊医嘱下达), , , , , mlng病人科室id)
    mstrPrivs = GetInsidePrivs(IIF(mint调用类型 = 1, pm门诊医嘱下达, pm住院医嘱下达))
    vs中药规格.Visible = True
    '初始化规格表格
    Call Grid.Init(vs中药规格, "规格,1305,1;产地,5070,1;剂量单位,900,1;数量,900,1;单价,900,1;药品ID")
    fra中药.Visible = True
    lblData.Visible = True
    cboData.Visible = True
    lblData.Caption = "煎法"
    lblNumZY.Visible = True
    cmdInsert.Visible = mbytUseType <> 3
    If mint场合 <> 3 Then
       lblZYStock.Visible = True
       Me.Height = Me.Height + lblZYStock.Height + 60
    End If
    If Not Init中药配方 Then Unload Me: Exit Sub

    '字体设置
    Call zlControl.SetPubFontSize(Me, mbytSize)
    '恢复个性化
    lngBaseHeight = Me.Height
    Call RestoreWinState(Me, App.ProductName, 2)
    
     '10.26.80增加规格显示后，以前个性化保存的高度可能不够
    If Me.Height < lngBaseHeight Then
        Me.Height = lngBaseHeight
    End If
    
    '窗体定位
    GetWindowRect mlngHwnd, vRect
    Me.Left = (vRect.Left - 1) * Screen.TwipsPerPixelX
    Me.Top = (vRect.Top - 1) * Screen.TwipsPerPixelY - Me.Height
    Call Form_Resize
    Call RefreshWeiNum
End Sub

Private Sub SetSplitLine()
'功能：设置中药配方输入的列分隔线
    Dim lngRow As Long, lngCol As Long
    Dim i As Long
        
    vsExt.Redraw = False
    lngRow = vsExt.Row: lngCol = vsExt.Col
    mblnChangeSel = True
    For i = 1 To mint每行味数
        vsExt.Select vsExt.FixedRows, i * 4 - 1, vsExt.Rows - 1, i * 4 - 1
        vsExt.CellBorder &HC0C0C0, 0, 0, 1, 0, 0, 0
    Next

    vsExt.ColWidth(0) = ((vsExt.Width - 60) / mint每行味数 - 285) * 0.45 '单味中药
    vsExt.ColWidth(1) = ((vsExt.Width - 60) / mint每行味数 - 285) * 0.22  '单味用量
    vsExt.ColWidth(2) = 285 '单位
    vsExt.ColWidth(3) = ((vsExt.Width - 60) / mint每行味数 - 285) * 0.33 '脚注
    For i = 4 To vsExt.Cols - 1
        vsExt.ColWidth(i) = vsExt.ColWidth(i - 4)
    Next
    
    vsExt.Row = lngRow: vsExt.Col = lngCol
    mblnChangeSel = False
    vsExt.Redraw = True
End Sub

Private Function Init中药配方() As Boolean
'功能：初始化中药配方表格格式及数据
'参数：mstrExtData=包含每味中药信息及煎法信息的串,为空时表示新输入中药配方
    Dim rsTmp As New ADODB.Recordset
    Dim rsTmpCopy As New ADODB.Recordset
    
    Dim strSQL As String, i As Long, j As Long
    Dim lngRow As Long, lngCol As Long, blnDo As Boolean
    Dim str中药IDs As String, lng煎法ID As Long, lngFirst药名ID As Long, lngFirst药品ID As Long
    Dim arr中药 As Variant, lng形态 As Long, dbl数量 As Double, str规格数量 As String
    Dim lngCur药名ID As Long, lngNext药名ID As Long, lng药品ID As Long
    Dim lng药房ID As Long, bln配方 As Boolean
    Dim strKey As String, blnSame散装中药 As Boolean
    Dim str名称 As String
    
    mstr配方明细 = ""
    vsExt.Clear
    vsExt.Cols = mint每行味数 * 4: vsExt.Rows = 7
  
    vsExt.FixedCols = 0: vsExt.FixedRows = 1
    vsExt.ColAlignment(0) = 1 '单味中药
    vsExt.ColAlignment(1) = 7 '单味用量
    vsExt.ColAlignment(2) = 1 '单位
    vsExt.ColAlignment(3) = 1 '脚注

    
    Me.Width = (Me.Width - Me.ScaleWidth) + IIF(mbytSize = 0, 2320, 2870) * mint每行味数 + 250
    Me.Height = Me.Height + vs中药规格.Height + fra中药.Height + 600

    For i = 4 To vsExt.Cols - 1
        vsExt.ColAlignment(i) = vsExt.ColAlignment(i - 4)
    Next
    vsExt.MergeCells = flexMergeFixedOnly
    vsExt.MergeRow(0) = True
    vsExt.Cell(flexcpAlignment, 0, 0, 0, vsExt.Cols - 1) = 1
    vsExt.Cell(flexcpText, 0, 0, 0, vsExt.Cols - 1) = "请先选择中药形态,然后依次输入中草药,单味用量,脚注。按*键选取中药或脚注。"
    vsExt.GridColor = vsExt.BackColor
    vsExt.Editable = flexEDKbdMouse
       
    vs中药规格.TabIndex = vsExt.TabIndex + 1
    txtJL.TabIndex = cboData.TabIndex + 1
    cmdOK.TabIndex = txtJL.TabIndex + 1

    On Error GoTo errH
    txt付数.Text = "1"
    txt付数.Tag = "1"
    If mint期效 = 0 Then    '长嘱不输付数
        txt付数.Enabled = False
        txt付数.BackColor = Me.BackColor
    End If
    
    If mstrExtData <> "" Then '修改
        lng煎法ID = Val(Split(mstrExtData, "|")(1))
        lng形态 = Val(Split(mstrExtData, "|")(2))
        txt付数.Text = Val(Split(mstrExtData, "|")(3))
        txtJL.Text = Split(mstrExtData, "|")(5)
        arr中药 = Split(Split(mstrExtData, "|")(0), ";")
        lng药品ID = Val(Split(arr中药(0), ",")(0))
                
        For i = 0 To UBound(arr中药)
            str中药IDs = str中药IDs & "," & CStr(Split(arr中药(i), ",")(0))
        Next
        str中药IDs = Mid(str中药IDs, 2)
                
        strSQL = "Select/*+ Rule*/ a.ID,b.药品ID,a.名称,a.计算单位,c.规格 as 规格 From 诊疗项目目录 A,药品规格 B,收费项目目录 C " & _
            "Where a.ID = b.药名ID And b.药品ID = C.ID And b.药品ID IN(Select Column_Value From Table(f_Num2list([1]))) "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str中药IDs)
        Set rsTmpCopy = zlDatabase.CopyNewRec(rsTmp)
        
        
        If vsExt.Rows < -Int(rsTmp.RecordCount / (-1 * mint每行味数)) + 1 Then
            vsExt.Rows = -Int(rsTmp.RecordCount / (-1 * mint每行味数)) + 1
        End If
        lngRow = vsExt.FixedRows: lngCol = 0
        
        '按照现在的内容和次序显示
        dbl数量 = 0
        str规格数量 = ""
        For i = 0 To UBound(arr中药)
            blnDo = True
            dbl数量 = dbl数量 + Val(Split(arr中药(i), ",")(1))
            str规格数量 = str规格数量 & ";" & Split(arr中药(i), ",")(0) & "," & Split(arr中药(i), ",")(1)
            If i < UBound(arr中药) Then
                rsTmp.Filter = "药品ID=" & CStr(Split(arr中药(i), ",")(0))
                lngCur药名ID = rsTmp!ID
                rsTmp.Filter = "药品ID=" & CStr(Split(arr中药(i + 1), ",")(0))
                lngNext药名ID = rsTmp!ID
                
                If lngCur药名ID = lngNext药名ID And lng形态 <> 0 Then
                    blnDo = False   '非散装的同种药的不同规格的数量累加
                End If
            End If
            
            If blnDo Then
                rsTmp.Filter = "药品ID=" & CStr(Split(arr中药(i), ",")(0))
                If Not rsTmp.EOF Then
                    str名称 = rsTmp!名称
                    If lng形态 = 0 Then '散装
                        strKey = rsTmp!ID & "_" & rsTmp!药品ID
                        
                        rsTmpCopy.Filter = "ID=" & rsTmp!ID
                        If rsTmpCopy.RecordCount > 0 Then
                            If rsTmpCopy.RecordCount > 1 Then blnSame散装中药 = True
                            If Not IsNull(rsTmp!规格) Then str名称 = str名称
                        End If
                    Else
                        strKey = "" & rsTmp!ID
                    End If
                    
                    str规格数量 = Mid(str规格数量, 2)
                    mcol规格数量.Add str规格数量, "_" & strKey
                    str规格数量 = ""
                
                    vsExt.TextMatrix(lngRow, lngCol) = str名称
                    vsExt.TextMatrix(lngRow, lngCol + 1) = FormatEx(dbl数量, 5): dbl数量 = 0
                    vsExt.TextMatrix(lngRow, lngCol + 2) = NVL(rsTmp!计算单位)
                    vsExt.TextMatrix(lngRow, lngCol + 3) = CStr(Split(arr中药(i), ",")(2))
                    
                    '用于恢复显示的记录
                    vsExt.Cell(flexcpData, lngRow, lngCol) = vsExt.TextMatrix(lngRow, lngCol)
                    vsExt.Cell(flexcpData, lngRow, lngCol + 1) = vsExt.TextMatrix(lngRow, lngCol + 1)
                    vsExt.Cell(flexcpData, lngRow, lngCol + 2) = strKey '记录中药ID
                    vsExt.Cell(flexcpData, lngRow, lngCol + 3) = vsExt.TextMatrix(lngRow, lngCol + 3)
                                    
                    '下一位置
                    If lngCol + 4 > vsExt.Cols - 1 Then
                        lngRow = lngRow + 1: lngCol = 0
                    Else
                        lngCol = lngCol + 4
                    End If
                End If
            End If
        Next
    Else '新增
        strSQL = "Select a.ID,a.类别,a.名称,a.计算单位 From 诊疗项目目录 a Where a.ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng项目ID)
        If rsTmp!类别 = "7" Then
            '输入了单味中草药
            vsExt.TextMatrix(vsExt.FixedRows, 0) = rsTmp!名称
            vsExt.TextMatrix(vsExt.FixedRows, 2) = NVL(rsTmp!计算单位)
            
            '用于恢复显示的记录
            vsExt.Cell(flexcpData, vsExt.FixedRows, 0) = vsExt.TextMatrix(vsExt.FixedRows, 0)
            lngFirst药名ID = CLng(rsTmp!ID)
            
            '长嘱按品种下达时，选择器返回的药品ID为0
            If mlng药品ID = 0 Then
                Set rsTmp = Get中药规格(lngFirst药名ID)
                If rsTmp.RecordCount > 0 Then
                    lng药品ID = rsTmp!药品ID
                    
                    '如果只有一种规格或一种形态，则取该规格的形态，否则缺省为散装形态
                    lng形态 = Val("" & rsTmp!中药形态)
                    rsTmp.Filter = "中药形态<>" & lng形态
                    If rsTmp.RecordCount > 1 Then lng形态 = 0
                Else
                    MsgBox "未找到该药品任何可用的规格，请选择其他药品", vbInformation, gstrSysName
                    Exit Function
                End If
            Else
                strSQL = "Select 中药形态 From 药品规格 Where 药品ID=[1]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng药品ID)
                lng形态 = Val("" & rsTmp!中药形态)
                lng药品ID = mlng药品ID
            End If
                        
            If lng形态 = 0 Then '散装
                strKey = lngFirst药名ID & "_" & lng药品ID
                mcol规格数量.Add lng药品ID & ",0", "_" & strKey
            Else
                strKey = lngFirst药名ID
                mcol规格数量.Add "", "_" & strKey
            End If
            If lng形态 = 0 Then '散装
                Set rsTmp = Get中药规格(lngFirst药名ID, lng形态)
                If rsTmp.RecordCount > 1 Then
                    rsTmp.Filter = "药品ID =" & lng药品ID
                    vsExt.TextMatrix(vsExt.FixedRows, 0) = rsTmp!名称
                    vsExt.Cell(flexcpData, vsExt.FixedRows, 0) = vsExt.TextMatrix(vsExt.FixedRows, 0)
                End If
            End If
            vsExt.Cell(flexcpData, vsExt.FixedRows, 2) = strKey '记录中药ID
        Else
            '输入了配方项目
            strSQL = "Select A.ID,A.名称,b.收费细目id as 药品id,A.计算单位,B.单次用量,B.医生嘱托,C.规格" & _
                " From 诊疗项目目录 A,诊疗项目组合 B,收费项目目录 C" & _
                " Where A.ID=B.诊疗项目ID And B.诊疗组合ID=[1] And c.Id(+) = b.收费细目id" & _
                " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
                " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL) And A.服务对象 IN([2],3) Order By B.序号"
                
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng项目ID, mint服务对象)
            If rsTmp.EOF Then
                MsgBox "该中药配方当前无有效的配方组成，请先到诊疗项目管理中设置。", vbInformation, gstrSysName
                Exit Function
            End If
            
            bln配方 = True
            If vsExt.Rows < -Int(rsTmp.RecordCount / (-1 * mint每行味数)) + 1 Then
                vsExt.Rows = -Int(rsTmp.RecordCount / (-1 * mint每行味数)) + 1
            End If
            lngRow = vsExt.FixedRows: lngCol = 0
            
            '按照设置的内容的次序显示
            For i = 1 To rsTmp.RecordCount
                vsExt.TextMatrix(lngRow, lngCol) = rsTmp!名称
                vsExt.TextMatrix(lngRow, lngCol + 1) = NVL(rsTmp!单次用量)
                vsExt.TextMatrix(lngRow, lngCol + 2) = NVL(rsTmp!计算单位)
                vsExt.TextMatrix(lngRow, lngCol + 3) = NVL(rsTmp!医生嘱托)
                
                '用于恢复显示的记录
                vsExt.Cell(flexcpData, lngRow, lngCol) = vsExt.TextMatrix(lngRow, lngCol)
                vsExt.Cell(flexcpData, lngRow, lngCol + 1) = vsExt.TextMatrix(lngRow, lngCol + 1)
                 '记录中药ID(对于散装且无药品ID，后面会重新设置为"药名Id_药品ID")
                vsExt.Cell(flexcpData, lngRow, lngCol + 2) = CLng(rsTmp!ID) & IIF(NVL(rsTmp!药品ID) = "", "", "_" & rsTmp!药品ID)
                vsExt.Cell(flexcpData, lngRow, lngCol + 3) = vsExt.TextMatrix(lngRow, lngCol + 3)
                
                If i = 1 Then
                    lngFirst药名ID = CLng(rsTmp!ID)
                    lngFirst药品ID = Val("" & rsTmp!药品ID)
                End If
                '下一位置
                If lngCol + 4 > vsExt.Cols - 1 Then
                    lngRow = lngRow + 1: lngCol = 0
                Else
                    lngCol = lngCol + 4
                End If
                rsTmp.MoveNext
            Next
            
            '获取配方项目的缺省煎法
            strSQL = "Select 用法ID From 诊疗用法用量 Where 性质=1 And 项目ID=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng项目ID)
            If Not rsTmp.EOF Then lng煎法ID = rsTmp!用法ID
            
            '获取第一味药的缺省规格(用来决定缺省形态和可用药房)
            Set rsTmp = Get中药规格(lngFirst药名ID, , , True)
            If rsTmp.RecordCount > 0 Then
                If lngFirst药品ID <> 0 Then
                    lng药品ID = lngFirst药品ID
                    rsTmp.Filter = "药品id=" & lngFirst药品ID
                    If Not rsTmp.EOF Then lng形态 = Val("" & rsTmp!中药形态)
                Else
                    lng药品ID = rsTmp!药品ID
                    
                    '如果只有一种规格或一种形态，则取该规格的形态，否则缺省为散装形态
                    lng形态 = Val("" & rsTmp!中药形态)
                    rsTmp.Filter = "中药形态<>" & lng形态
                    If rsTmp.RecordCount > 1 Then lng形态 = 0
                End If
            End If
        End If
    End If
    vsExt.ScrollBars = flexScrollBarNone
    
    If mint调用类型 = 2 Then
        strSQL = "select a.病人性质 from 病案主页 a where a.病人id=[1] and a.主页id=[2]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID, Val(mvar就诊ID))
        If Not rsTmp.EOF Then mlng病人性质 = Val(rsTmp!病人性质 & "")
    End If
        
    '中药煎法
    strSQL = "Select A.ID,A.编码,A.名称 From 诊疗项目目录 A" & _
        " Where A.类别='E' And A.操作类型='3'" & IIF(mlng病人性质 = 1, "", " And A.服务对象 IN([1],3)") & _
        " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
        " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
        IIF(mint场合 <> 3 And mlng病人性质 <> 1, " And (Exists(Select 1 From 诊疗适用科室 Where 项目ID=A.ID And 科室ID=[2])" & _
                            " Or Not Exists(Select 1 From 诊疗适用科室 Where 项目ID=A.ID))", "") & _
        " Order by A.编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mint服务对象, mlng病人科室id)
    If rsTmp.Filter <> 0 Then rsTmp.Filter = 0
    If rsTmp.EOF Then
        MsgBox "未找到有效的中药煎法，请先到诊疗项目管理中设置。", vbInformation, gstrSysName
        Exit Function
    End If
    
    For i = 1 To rsTmp.RecordCount
        cboData.AddItem rsTmp!编码 & "-" & rsTmp!名称
        cboData.ItemData(cboData.NewIndex) = rsTmp!ID
        If rsTmp!ID = lng煎法ID Then
            Call Cbo.SetIndex(cboData.hwnd, cboData.NewIndex)
        End If
        rsTmp.MoveNext
    Next
    If cboData.ListCount = 1 And cboData.ListIndex = -1 Then Call Cbo.SetIndex(cboData.hwnd, 0)
    
    '可用药房(当指定药房限定库存时，受药品ID的影响)
    Call Get中药房(cbo药房, lng药品ID, mlng病人科室id, mint服务对象, mlng中药房)
    If mstrExtData <> "" Then
        lng药房ID = Val(Split(mstrExtData, "|")(4))
    Else
        lng药房ID = IIF(mlngPreRow中药房 = 0, mlng中药房, mlngPreRow中药房)
    End If
    Call Cbo.Locate(cbo药房, lng药房ID, True)
    If cbo药房.ListCount > 0 And cbo药房.ListIndex = -1 Then Call Cbo.SetIndex(cbo药房.hwnd, 0)
    
    
    '中药形态
    arr中药 = Split("散装(&0),饮片(&1),免煎剂(&2)", ",")
    For i = 0 To 2
        optMode(i).Visible = True
        optMode(i).Enabled = True
        optMode(i).Caption = arr中药(i)
        optMode(i).Width = optMode(i).Width + IIF(i = 2, 450, 300)
        If i = lng形态 Then optMode(i).value = True
    Next
    If blnSame散装中药 Then
        optMode(1).Enabled = False
        optMode(2).Enabled = False
    End If
    Call SetSameItem
    If mstrExtData = "" And bln配方 Then
        '输入定制的“中药配方”时，数量已预设
        For i = vsExt.FixedRows To vsExt.Rows - 1
            For j = 0 To vsExt.Cols - 1 Step 4
                strKey = vsExt.Cell(flexcpData, i, j + 2)
                
                lngCur药名ID = 0 '清空lngCur药名ID
                
                If strKey <> "" Then lngCur药名ID = Val(Split(strKey, "_")(0))
                
                If lngCur药名ID <> 0 Then
                    dbl数量 = Val(vsExt.TextMatrix(i, j + 1))
                    '没有设置规格的药重新取规格
                    str规格数量 = ""
                    If InStr(strKey, "_") = 0 Then
                        '缺省规格
                        Set rsTmp = Get中药规格(lngCur药名ID, lng形态)
                                             
                        strKey = lngCur药名ID
                        If rsTmp.RecordCount > 0 Then
                            If lng形态 = 0 Then '散装
                                strKey = lngCur药名ID & "_" & rsTmp!药品ID
                                vsExt.Cell(flexcpData, i, j + 2) = strKey
                            End If
                            str规格数量 = rsTmp!药品ID & "," & 0
                        End If
                    ElseIf cbo药房.ListIndex <> -1 Then
                        Set rsTmp = Get药品规格(Val(Split(strKey, "_")(1)), True)
                        If rsTmp.RecordCount > 0 Then str规格数量 = Val(Split(strKey, "_")(1)) & "," & 0
                    End If
                    
                    mcol规格数量.Add str规格数量, "_" & strKey
                    
                    Call Split中药规格(lngCur药名ID, dbl数量, strKey)
                    
                    If mcol规格数量("_" & strKey) = "" Or InStr(mcol规格数量("_" & strKey), "|") > 0 Then
                        vsExt.Cell(flexcpForeColor, i, j + 1) = vbRed
                    End If
                End If
                If Val(vsExt.TextMatrix(i, j + 1)) <> 0 Then
                    mstr配方明细 = mstr配方明细 & ";" & Replace(CStr(mcol规格数量("_" & strKey)), ";", "," & vsExt.TextMatrix(i, j + 3) & ";") & "," & vsExt.TextMatrix(i, j + 3)
                End If
            Next
        Next
    End If
    
    vsExt.Row = vsExt.FixedRows: vsExt.Col = 1
    Init中药配方 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub RefreshWeiNum()
'功能：确认味数
    Dim intNum As Integer
    Dim i As Long, j As Long
    
    intNum = 0
    For i = vsExt.FixedRows To vsExt.Rows - 1
        For j = vsExt.FixedCols To vsExt.Cols - 1 Step 4
            If vsExt.TextMatrix(i, j) <> "" And vsExt.Cell(flexcpData, i, j) <> 0 Then intNum = intNum + 1
        Next
    Next
    lblNumZY = "共 " & intNum & " 味"
End Sub


Private Function Get中药规格(ByVal lng药名ID As Long, Optional ByVal lng形态 As Long = -1, Optional ByVal blnFirst As Boolean, Optional ByVal bln配方 As Boolean) As ADODB.Recordset
'功能：根据中药诊疗ID获取中药规格
'参数：bln配方 true=新开时调用配方
    Dim strSQL As String, lng药房ID As Long
    
    On Error GoTo errH
    If lng形态 = 0 Then
        Set Get中药规格 = Get药品规格(lng药名ID)
    Else
        If mstr可用中药房 <> "" Then
            If gblnStock And Not blnFirst Then
                If cbo药房.ListIndex = -1 Then
                    lng药房ID = IIF(mlngPreRow中药房 = 0, mlng中药房, mlngPreRow中药房)
                    If bln配方 And mlngPreRow中药房 <> 0 Then lng药房ID = 0
                Else
                    lng药房ID = cbo药房.ItemData(cbo药房.ListIndex)
                End If
                '如果是成套调用，药房ID=0则不加库存条件
                If mint场合 <> 3 Or lng药房ID <> 0 Then
                    strSQL = " And Exists(Select 1 From 药品库存 B" & _
                        " Where (Nvl(b.批次, 0) = 0 Or b.效期 Is Null Or b.效期>Trunc(Sysdate))" & _
                        " And b.性质=1 And a.药品ID=b.药品ID" & IIF(lng药房ID = 0, "", " And b.库房ID=[2]") & _
                        " And b.可用数量>0)"
                End If
            End If
        End If
    
        strSQL = "Select A.药品ID,A.中药形态,D.编码 From 药品规格 A,收费项目目录 D Where A.药名ID = [1] And A.药品ID = D.ID" & _
             IIF(lng形态 = -1, "", " And A.中药形态 = [4]") & strSQL & _
             " And (D.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or D.撤档时间 IS NULL) And D.服务对象 IN([3],3)" & _
             " And (D.站点='" & gstrNodeNo & "' Or D.站点 is Null) Order by D.编码"
        Set Get中药规格 = zlDatabase.OpenSQLRecord(strSQL, "读取中药规格", lng药名ID, lng药房ID, mint服务对象, lng形态)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Function Get中药房(objCbo As ComboBox, ByVal lng药品ID As Long, ByVal lng病人科室ID As Long, _
    ByVal int范围 As Integer, Optional ByVal lng当前药房ID As Long) As Boolean
'功能：读取可用中药房，并加载到下拉列表中
'参数：
'      int范围=1-门诊,2-住院(缺省)
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim bytDay As Byte, bln上班安排 As Boolean
    Dim bln规格 As Boolean, i As Long
    Dim strStock As String
        
   
    '药品库存限制
    If mstr可用中药房 <> "" Then
        If gblnStock Then
            strStock = " And Exists(" & _
                " Select 1 From 药品库存" & _
                " Where (Nvl(批次,0)=0 Or 效期 Is Null Or 效期>Trunc(Sysdate))" & _
                " And 性质=1 And 药品ID=[3] And 库房ID=A.执行科室ID" & _
                " And 可用数量>0 And Instr('," & mstr可用中药房 & ",',','||库房ID||',')>0)"
        Else
            strStock = " And Instr('," & mstr可用中药房 & ",',','||A.执行科室ID||',')>0"
        End If
    End If
          
     '药品从系统指定的储备药房中找
    If mint场合 <> 3 Then
        If int范围 = 1 Then bln上班安排 = Check上班安排() '住院医嘱不管药房上班安排
    
        If bln上班安排 Then
            bytDay = Weekday(zlDatabase.Currentdate, vbMonday) Mod 7 '0=周日,1=周一
        End If
    End If
    strSQL = _
         " Select Distinct C.ID,C.编码,C.简码,C.名称,B.服务对象" & _
         " From 收费执行科室 A,部门性质说明 B,部门表 C" & IIF(bln上班安排, ",部门安排 D", "") & _
         " Where A.执行科室ID+0=B.部门ID And B.工作性质='中药房' And B.服务对象 IN([1],3) And B.部门ID=C.ID" & _
         " And (A.病人来源 is NULL Or A.病人来源=[1]) " & IIF(mint场合 <> 3, " And (A.开单科室ID is NULL Or A.开单科室ID=[2])", "") & _
         " And A.收费细目ID=[3]" & strStock & _
         " And (C.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or C.撤档时间 is NULL) And (C.站点='" & gstrNodeNo & "' Or C.站点 is Null)" & _
         IIF(bln上班安排, " And D.部门ID=C.ID And D.星期=[4] And To_Char(Sysdate,'HH24:MI:SS') Between To_Char(D.开始时间,'HH24:MI:SS') and To_Char(D.终止时间,'HH24:MI:SS')", "") & _
         " Order by B.服务对象,C.编码"
     
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "读取中药房", int范围, lng病人科室ID, lng药品ID, bytDay)
    objCbo.Clear
    For i = 1 To rsTmp.RecordCount
        objCbo.AddItem rsTmp!编码 & "-" & rsTmp!名称
        objCbo.ItemData(i - 1) = Val(rsTmp!ID)
        If lng当前药房ID = Val(rsTmp!ID) Then
            Call Cbo.SetIndex(objCbo.hwnd, i - 1)
        End If
        rsTmp.MoveNext
    Next
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SetSameItem(Optional ByVal lng诊疗项目ID As Long)
'功能：散装中药是否界面上存在多个规格使用,存在则显示规格名称,否则取消
    Dim i As Long, j As Long, strKey As String
    Dim strItem As String, strTmp As String
    Dim arrTmp As Variant
    Dim rsTmp As Recordset, strSQL As String
    Dim str中药IDs As String
    Dim lngRow As Long, lngCol As Long
    Dim lng药名ID As Long, lng药品ID As Long
    
    If Get中药形态 <> 0 Then Exit Sub
    For i = vsExt.FixedRows To vsExt.Rows - 1
        For j = 0 To vsExt.Cols - 1 Step 4
            strKey = vsExt.Cell(flexcpData, i, j + 2)
            If strKey <> "" And InStr(strKey, "_") > 0 Then
                If lng诊疗项目ID = 0 Then
                    strItem = strItem & "," & strKey & "|" & i & "|" & j
                    str中药IDs = str中药IDs & "," & Val(Mid(strKey, InStr(strKey, "_") + 1))
                ElseIf Val(Mid(strKey, 1, InStr(strKey, "_") - 1)) = lng诊疗项目ID Then
                    strItem = strItem & "," & strKey & "|" & i & "|" & j
                    str中药IDs = str中药IDs & "," & Val(Mid(strKey, InStr(strKey, "_") + 1))
                End If
            End If
        Next
    Next
    strItem = Mid(strItem, 2)
    str中药IDs = Mid(str中药IDs, 2)
    If strItem = "" Then Exit Sub
    
    strSQL = "Select b.药名id,b.药品id,c.名称,a.规格 From 收费项目目录 a,药品规格 b,诊疗项目目录 c" & _
        " Where a.id=b.药品id and c.id=b.药名id and a.ID IN (Select Column_Value From Table(f_Num2list([1])))"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str中药IDs)

    arrTmp = Split(strItem, ",")
    For i = 0 To UBound(arrTmp)
        strTmp = arrTmp(i)
        strKey = Split(arrTmp(i), "|")(0)
        lngRow = Val(Split(arrTmp(i), "|")(1))
        lngCol = Val(Split(arrTmp(i), "|")(2))
        
        If InStr(strKey, "_") > 0 Then
            lng药名ID = Val(Split(strKey, "_")(0))
            lng药品ID = Val(Split(strKey, "_")(1))
        Else
            lng药名ID = Val(strKey)
            lng药品ID = 0
        End If
        
        rsTmp.Filter = "药名id=" & lng药名ID
        If Not rsTmp.EOF Then
            vsExt.TextMatrix(lngRow, lngCol) = rsTmp!名称 & ""
        End If
        
        If rsTmp.RecordCount > 1 Then
            rsTmp.Filter = "药名id=" & lng药名ID & " and 药品ID=" & lng药品ID
            If Not rsTmp.EOF Then
                vsExt.TextMatrix(lngRow, lngCol) = rsTmp!名称 & "(" & rsTmp!规格 & ")"
            End If
        End If
        vsExt.Cell(flexcpData, lngRow, lngCol) = vsExt.TextMatrix(lngRow, lngCol)
    Next
End Sub

Private Function Get药品规格(ByVal lng药名ID As Long, Optional ByVal bln规格 As Boolean, Optional ByVal lng形态 As Long) As ADODB.Recordset
'功能：获取当前药品的指定形态的可用的规格
'参数：
'      bln规格 当该参数传入为true时，lng药名ID 就当作 药品id 来使用
'      int形态=0-散装;1-饮片；2-免煎剂
    Dim lng药房ID As Long, strSQL As String
    
    If mstr可用中药房 <> "" Then
        If gblnStock Then
            If cbo药房.ListIndex <> -1 Then lng药房ID = cbo药房.ItemData(cbo药房.ListIndex)
            If mint场合 <> 3 Or lng药房ID <> 0 Then
                strSQL = " And Exists(Select 1 From 药品库存 B" & _
                    " Where (Nvl(b.批次, 0) = 0 Or b.效期 Is Null Or b.效期>Trunc(Sysdate))" & _
                    " And b.性质=1 And a.药品ID=b.药品ID" & IIF(lng药房ID = 0, "", " And b.库房ID=[2]") & _
                    " And b.可用数量>0)"
            End If
        End If
    End If
    
    strSQL = "Select a.药名id, a.药品id, d.规格, d.产地, a.剂量系数, d.编码, d.名称,A.中药形态,d.是否变价" & vbNewLine & _
            "From 药品规格 A, 收费项目目录 D" & vbNewLine & _
            "Where a.中药形态 = [4] And a.药品ID = d.ID" & strSQL & vbNewLine & _
            IIF(bln规格, " And a.药品id = [1]", " And a.药名id = [1]") & vbNewLine & _
            " And (D.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or D.撤档时间 IS NULL) And D.服务对象 IN([3],3)" & _
            " And (D.站点='" & gstrNodeNo & "' Or D.站点 is Null) Order By D.编码"
    On Error GoTo errH
    Set Get药品规格 = zlDatabase.OpenSQLRecord(strSQL, "规格列表", lng药名ID, lng药房ID, mint服务对象, lng形态)
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Split中药规格(ByVal lng药名ID As Long, ByVal dbl数量 As Double, ByVal strKey As String, Optional ByVal str药品IDs As String)
'功能：输入数量后，进行规格数量的分配(存储到mcol规格数量中)
'参数：strKey=散装：药名ID_药品ID，非散装:药品ID
'      str药品IDs=当前使用的药品ID字符串
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim str规格数量 As String, lng药房ID As Long, lng形态 As Long
    Dim lng药品ID As Long
        
    If mblnSelf = True Then Exit Sub
    If InStr(strKey, "_") > 0 Then
        lng形态 = 0
    Else
        lng形态 = Get中药形态
    End If
    
    If lng形态 = 0 Then
        '散装的在输入时已确定规格
        str规格数量 = mcol规格数量("_" & strKey)
        If str规格数量 <> "" Then str规格数量 = Split(str规格数量, ",")(0) & "," & FormatEx(dbl数量, 5)
    Else
        '2.分配结果,药品id,数量;药品id,数量;...|剩余数量
        On Error GoTo errH
        If cbo药房.ListIndex <> -1 Then
            lng药房ID = cbo药房.ItemData(cbo药房.ListIndex)
        End If
        '如果没有药房，则先重设规格后再重新分配
        If lng药房ID = 0 Then
            '防止死循环
            mblnSelf = True
            Call ReSet中药规格
            If cbo药房.ListIndex = -1 And mint场合 <> 3 Then
                lng药品ID = 0
                strKey = vsExt.Cell(flexcpData, 1, 2)
                If InStr(strKey, "_") > 0 Then
                    lng药品ID = Val(Mid(strKey, InStr(strKey, "_") + 1))
                Else
                    Set rsTmp = Get中药规格(Val(strKey), Get中药形态)
                    If rsTmp.RecordCount > 0 Then
                        lng药品ID = Val(rsTmp!药品ID & "")
                    End If
                End If
                lng药房ID = IIF(mlngPreRow中药房 = 0, mlng中药房, mlngPreRow中药房)
                Call Get中药房(cbo药房, lng药品ID, mlng病人科室id, mint服务对象, lng药房ID)
                lng药房ID = 0
                If cbo药房.ListIndex <> -1 Then
                    If cbo药房.ListCount > 0 Then cbo药房.ListIndex = 0
                    lng药房ID = cbo药房.ItemData(cbo药房.ListIndex)
                End If
            End If
            mblnSelf = False
        End If
        strSQL = "Select Zl_Dispensechspecs([1],[2],[3],[4],[5],Null,[6],[7]) as txt From dual"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "规格分配", lng药名ID, lng形态, dbl数量, Val(txt付数.Text), lng药房ID, mint调用类型, str药品IDs)
        str规格数量 = "" & rsTmp!txt
    End If
    
    Call mcol规格数量.Remove("_" & strKey)
    mcol规格数量.Add str规格数量, "_" & strKey
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function Check上班安排() As Boolean
'功能：检查中药房是否启用了上班安排
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    Static bln药房Load As Boolean
    Static bln药房Last As Boolean
    
    '是否有安排只需读取一次
    If bln药房Load Then Check上班安排 = bln药房Last: Exit Function
     
    On Error GoTo errH
    strSQL = "Select Count(B.部门ID) as NUM From 部门性质说明 A,部门安排 B" & _
            " Where A.部门ID=B.部门ID And A.工作性质 ='中药房'"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, "Check上班安排")
    If Not rsTmp.EOF Then
        Check上班安排 = NVL(rsTmp!Num, 0) > 0
    End If
    
    bln药房Load = True: bln药房Last = Check上班安排
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub ReSet中药规格(Optional ByVal blnReset As Boolean = True)
'功能：重设所有中药的规格(按数量分配),并重新显示当前中药的数量按规格分配的列表
'参数：blnReset=false 不重设规格
    Dim i As Long, j As Long, lng药名ID As Long, dbl数量 As Double, lng药品ID As Long
    Dim lng形态 As Long, rsTmp As ADODB.Recordset
    Dim strKey As String, strTmp As String
    
    lng形态 = Get中药形态
    For i = vsExt.FixedRows To vsExt.Rows - 1
        For j = 0 To vsExt.Cols - 1 Step 4
            strKey = vsExt.Cell(flexcpData, i, j + 2)
            If strKey <> "" Then lng药名ID = Val(Split(strKey, "_")(0))
            
            If strKey <> "" Then
                strTmp = Get药品名称(lng药名ID)
                If lng形态 = 0 Then
                    If mcol规格数量("_" & strKey) <> "" Then lng药品ID = Val(Split(mcol规格数量("_" & strKey), ";")(0))
                    mcol规格数量.Remove ("_" & strKey)  '重选规格ID
                    If InStr(strKey, "_") > 0 And blnReset = False Then
                        Set rsTmp = Get药品规格(Val(Mid(strKey, InStr(strKey, "_") + 1)), True) '取原规格
                        lng药品ID = Val(Mid(strKey, InStr(strKey, "_") + 1))
                        strTmp = vsExt.TextMatrix(i, j)
                    Else
                        Set rsTmp = Get药品规格(lng药名ID)  '取缺省规格
                    End If
                    rsTmp.Filter = "药品id=" & lng药品ID    '如果原规格可用，则保持不变
                    If rsTmp.RecordCount = 0 Then rsTmp.Filter = ""
                    
                    If rsTmp.RecordCount > 0 Then
                        strKey = rsTmp!药名ID & "_" & rsTmp!药品ID
                        mcol规格数量.Add rsTmp!药品ID & ",0", "_" & strKey
                    Else
                        If lng药品ID = 0 Then
                            strKey = lng药名ID
                        Else
                            strKey = lng药名ID & "_" & lng药品ID
                        End If
                        mcol规格数量.Add "", "_" & strKey
                    End If
                Else
                    mcol规格数量.Remove ("_" & strKey) '以前可能是散装,Key可能由"药名ID_药品ID"变为"药品ID"，所以要先删除
                    strKey = lng药名ID
                    mcol规格数量.Add "", "_" & strKey
                End If
                vsExt.Cell(flexcpData, i, j + 2) = strKey
                vsExt.TextMatrix(i, j) = strTmp
                vsExt.Cell(flexcpData, i, j) = strTmp
                dbl数量 = Val(vsExt.TextMatrix(i, j + 1))
                If dbl数量 <> 0 Then Call Split中药规格(lng药名ID, dbl数量, strKey)
                
                If mcol规格数量("_" & strKey) = "" Or InStr(mcol规格数量("_" & strKey), "|") > 0 Then
                    vsExt.Cell(flexcpForeColor, i, j + 1) = vbRed
                Else
                    vsExt.Cell(flexcpForeColor, i, j + 1) = vsExt.ForeColor
                End If
            End If
        Next
    Next
    strKey = vsExt.Cell(flexcpData, vsExt.Row, (vsExt.Col \ 4) * 4 + 2)
    
    lng药名ID = 0 '清空lng药名ID
    
    If strKey <> "" Then lng药名ID = Val(Split(strKey, "_")(0))
    Call SetSameItem
    
    If lng药名ID <> 0 Then
        dbl数量 = Val(vsExt.TextMatrix(vsExt.Row, (vsExt.Col \ 4) * 4 + 1))
        Call Show中药规格(lng药名ID, dbl数量)
    End If
End Sub

Private Function Get中药形态() As Long
    Dim i As Long
    
    For i = 0 To optMode.UBound
        If optMode(i).value = True Then Exit For
    Next
    Get中药形态 = i
End Function

Private Function Get药品名称(ByVal lng药名ID As Long) As String
'功能：返回药品名称
    On Error GoTo errH
    
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
    strSQL = "Select 名称 From 诊疗项目目录 Where ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "mdlCISKernel", lng药名ID)
    If Not rsTmp.EOF Then Get药品名称 = rsTmp!名称
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Show中药规格(ByVal lng药名ID As Long, ByVal dbl数量 As Double, Optional ByVal lng形态 As Long = -1)
'功能：根据当前行和列，显示或隐藏中药规格列表
'      如果是散装形态，则加载可选择的规格下拉列表

    Dim str规格数量 As String, arrTmp As Variant, arrValue As Variant
    Dim i As Long, str药品IDs As String, lngColBegin As Long
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim strMsg As String, strKey As String
    
    lngColBegin = (vsExt.Col \ 4) * 4
    cmd形态.Visible = False
        
    With vs中药规格
        .Rows = .FixedRows
        .ColComboList(col规格) = ""
        '成套编辑不显示库存
        If mint场合 <> 3 Then
            '显示中药库存
            strSQL = "Select d.编码, d.规格, d.名称, e.名称 As 药房, d.计算单位,Sum(Nvl(m.可用数量, 0))  As 可用数量" & vbNewLine & _
                    "From 药品库存 M, 药品规格 A, 收费项目目录 D, 部门表 E" & vbNewLine & _
                    "Where m.药品id = d.Id And m.药品id = a.药品id And m.库房id = e.Id And" & vbNewLine & _
                    "      (Nvl(m.批次, 0) = 0 Or m.效期 Is Null Or m.效期 > Trunc(Sysdate)) And a.药名id = [1] And m.库房id = [2] And" & vbNewLine & _
                    "      (d.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or d.撤档时间 Is Null) And d.服务对象 In ([3],'3') And" & vbNewLine & _
                    "      (d.站点 = '" & gstrNodeNo & "' Or d.站点 Is Null)" & vbNewLine & _
                    "Group By e.名称, d.编码, d.规格, d.名称, d.计算单位" & vbNewLine & _
                    "Having Sum(Nvl(m.可用数量, 0)) > 0" & vbNewLine & _
                    "Order By d.编码"
            lblZYStock.Caption = ""
            If cbo药房.ListIndex <> -1 Then
                If lng药名ID = 0 Then Exit Sub
                On Error GoTo errH
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng药名ID, cbo药房.ItemData(cbo药房.ListIndex), mint服务对象 & "")
                If rsTmp.RecordCount > 0 Then
                    If InStr(mstrPrivs, "显示药品库存") = 0 Then
                        lblZYStock.Caption = "库存：有。"
                    Else
                        Do While Not rsTmp.EOF
                            lblZYStock.Caption = lblZYStock.Caption & IIF(lblZYStock.Caption = "", "库存：", "    ") & rsTmp!规格 & ":" & rsTmp!可用数量 & rsTmp!计算单位
                            rsTmp.MoveNext
                        Loop
                    End If
                Else
                    lblZYStock.Caption = "库存：无。"
                End If
            End If
        End If
        If lng形态 = -1 Then lng形态 = Get中药形态
        
        If dbl数量 = 0 And lng形态 <> 0 Or cbo药房.ListIndex = -1 And mint场合 <> 3 Then Exit Sub
        vsExt.Cell(flexcpForeColor, vsExt.Row, lngColBegin + 1) = vsExt.ForeColor
        
        strKey = vsExt.Cell(flexcpData, vsExt.Row, (vsExt.Col \ 4) * 4 + 2)
        If strKey = "" And lng药名ID <> 0 Then strKey = lng药名ID
        If strKey <> "" Then str规格数量 = Trim(mcol规格数量("_" & strKey))

        .Redraw = False
        If str规格数量 = "" Then
            .Rows = .FixedRows + 1
            '不能分配时返回空,例如:剂量为6和10的情况下,3克的分配
            .MergeCells = flexMergeRestrictRows
            If lng形态 = 0 Then
                strMsg = "该药品没有可用的散装形态，请选择其它药品或形态。"
            Else
                strMsg = "无法将所有数量按可用规格分配，请调整用量。"
            End If
            
            .TextMatrix(.Rows - 1, col规格) = strMsg
            .TextMatrix(.Rows - 1, col产地) = strMsg
            .MergeRow(.Rows - 1) = True
            .TextMatrix(.Rows - 1, col数量) = FormatEx(dbl数量, 5)
            .Cell(flexcpForeColor, .Rows - 1, col数量) = vbRed
            
            vsExt.Cell(flexcpForeColor, vsExt.Row, lngColBegin + 1) = vbRed
            If lng形态 <> 0 Then
                '散装形态及饮片免煎也允许选择规格（66074）
                Set rsTmp = Get药品规格(lng药名ID, , lng形态)
                If rsTmp.RecordCount > 1 Then
                    str药品IDs = ""
                    For i = 1 To rsTmp.RecordCount
                        str药品IDs = str药品IDs & "|#" & rsTmp!药品ID & ";" & rsTmp!编码 & "-" & rsTmp!名称 & IIF(Not IsNull(rsTmp!规格), "(" & rsTmp!规格 & ")", "")
                        rsTmp.MoveNext
                    Next
                    .ColComboList(col规格) = Mid(str药品IDs, 2)
                    rsTmp.MoveFirst
                    .RowData(.FixedRows) = rsTmp   '只有一行
                    .Cell(flexcpBackColor, .FixedRows, col规格, .Rows - 1, col规格) = &HF0F4E4
                End If
            End If
        Else
            arrTmp = Split(Split(str规格数量, "|")(0), ";")
            
            If InStr(str规格数量, "|") > 0 Then
                vs中药规格.Rows = vs中药规格.FixedRows + UBound(arrTmp) + 2
            Else
                vs中药规格.Rows = vs中药规格.FixedRows + UBound(arrTmp) + 1
            End If
            For i = 0 To UBound(arrTmp)
                arrValue = Split(arrTmp(i), ",")
                str药品IDs = str药品IDs & "," & Val(arrValue(0))
                .TextMatrix(.FixedRows + i, col药品ID) = Val(arrValue(0))  '规格ID
                .TextMatrix(.FixedRows + i, col数量) = FormatEx(arrValue(1), 5)    '数量
            Next
            str药品IDs = Mid(str药品IDs, 2)
            
            On Error GoTo errH
            '读出所有可用(有库存)的散装规格，以便可以选择其它的规格
             '散装形态及饮片免煎也允许选择规格（66074）
            Set rsTmp = Get药品规格(lng药名ID, , lng形态)
            For i = .FixedRows To .Rows - 1
                If InStr(str规格数量, "|") > 0 And i = .Rows - 1 Then
                '最后一行显示未分配数量
                    .MergeCells = flexMergeRestrictRows
                    strMsg = "无法将所有数量按可用规格分配，请调整用量。"
                    .TextMatrix(i, col规格) = strMsg
                    .TextMatrix(i, col产地) = strMsg
                    .MergeRow(i) = True
                    .Cell(flexcpForeColor, i, col数量) = vbRed
                    .TextMatrix(i, col数量) = FormatEx(Split(str规格数量, "|")(1), 5)
                    vsExt.Cell(flexcpForeColor, vsExt.Row, lngColBegin + 1) = vbRed
                Else
                    rsTmp.Filter = "药品ID = " & CStr(.TextMatrix(i, col药品ID))
                    If rsTmp.RecordCount = 0 Then '散装，库存不足时（允许保存）
                        strMsg = "当前药房库存不足，或者没有散装规格。"
                        .TextMatrix(.Rows - 1, col规格) = strMsg
                        .TextMatrix(.Rows - 1, col产地) = strMsg
                        .MergeRow(.Rows - 1) = True
                    
                        .Cell(flexcpForeColor, i, col数量) = vbRed
                        vsExt.Cell(flexcpForeColor, vsExt.Row, lngColBegin + 1) = vbRed
                    Else
                        .TextMatrix(i, col规格) = "" & rsTmp!规格
                        .Cell(flexcpData, i, col规格) = "" & rsTmp!规格 '用于散装规格取消下拉选择时恢复
                        .TextMatrix(i, col产地) = "" & rsTmp!产地
                        .TextMatrix(i, col剂量单位) = vsExt.TextMatrix(vsExt.Row, lngColBegin + 2)
                        
                        '记录售价单价
                        If NVL(rsTmp!是否变价, 0) = 0 Then
                            .TextMatrix(i, col单价) = Format(CalcPrice(Val("" & rsTmp!药品ID), , , True, , , mstr药品价格等级), gstrDecPrice)
                        Else '时价
                            .TextMatrix(i, col单价) = Format(CalcDrugPrice(Val("" & rsTmp!药品ID), cbo药房.ItemData(cbo药房.ListIndex), Val(.TextMatrix(i, col数量)), , True, 1, mstr药品价格等级), gstrDecPrice)
                        End If
                    End If
                End If
            Next
            
            '散装形态及饮片免煎也允许选择规格（66074）
            rsTmp.Filter = ""
            If rsTmp.RecordCount > 1 Then
                str药品IDs = ""
                For i = 1 To rsTmp.RecordCount
                    str药品IDs = str药品IDs & "|#" & rsTmp!药品ID & ";" & rsTmp!编码 & "-" & rsTmp!名称 & IIF(Not IsNull(rsTmp!规格), "(" & rsTmp!规格 & ")", "")
                    rsTmp.MoveNext
                Next
                .ColComboList(col规格) = Mid(str药品IDs, 2)
                rsTmp.MoveFirst
                .RowData(.FixedRows) = rsTmp   '只有一行
                .Cell(flexcpBackColor, .FixedRows, col规格, .Rows - 1, col规格) = &HF0F4E4
            End If
        End If
        
        If lng形态 <> 0 Then
            '非散装形态，未分配完时，允许换为散装
            If str规格数量 = "" Or InStr(str规格数量, "|") > 0 Then
                Set rsTmp = Get药品规格(lng药名ID)
                If rsTmp.RecordCount > 0 Then
                    strMsg = "无法将所有数量按可用规格分配，请调整用量或改用散装。"
                    .TextMatrix(.Rows - 1, col规格) = strMsg
                    .TextMatrix(.Rows - 1, col产地) = strMsg
                    .MergeRow(.Rows - 1) = True
                    
                    .Select .Rows - 1, col剂量单位
                    cmd形态.Visible = True
                    cmd形态.Tag = rsTmp!药品ID  '缺省规格
                    cmd形态.Caption = "散装(&D)"
                    cmd形态.Top = vs中药规格.CellTop
                    cmd形态.Left = vs中药规格.CellLeft
                    cmd形态.Width = vs中药规格.CellWidth
                    cmd形态.Height = vs中药规格.CellHeight
                End If
            End If
        End If
        
        .Redraw = True
    End With
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName, 2)

    mlngHwnd = 0
    mbln医保 = False
    mint险类 = 0
    mint婴儿 = 0
    mlng病人ID = 0
    mlng病人科室id = 0
    mvar就诊ID = Empty
    mstr性别 = ""
    mint场合 = 0
    mbytUseType = 0
    mint期效 = 0
    mint服务对象 = 0
    mint调用类型 = 0
    mlng项目ID = 0
    mlng药品ID = 0
    mlngPreRow中药房 = 0
    Set mclsInsure = Nothing
    Set mcol规格数量 = Nothing
    Set mfrmParent = Nothing
End Sub

Private Sub optMode_Click(Index As Integer)
    Dim lng药名ID As Long, lng药房ID As Long, str规格数量 As String, lng药品ID As Long
    Dim rsTmp As ADODB.Recordset
    Dim strKey As String, bln重置药房 As Boolean
    
    If Not Me.Visible Then Exit Sub
    strKey = vsExt.Cell(flexcpData, vsExt.FixedRows, vsExt.FixedCols + 2)
    If strKey <> "" Then lng药名ID = Val(Split(strKey, "_")(0))
    
    If lng药名ID <> 0 And gblnStock Then     '指定药房时限定库存时，形态改变后，第一味药的缺省规格可能变了，可用药房就变了
        
        str规格数量 = mcol规格数量("_" & strKey)
        If str规格数量 <> "" Then
            Set rsTmp = Get中药规格(lng药名ID, Index)
            If rsTmp.RecordCount > 0 Then
                lng药品ID = Val(Split(str规格数量, ",")(0))
                If lng药品ID <> Val(rsTmp!药品ID) Then
                    rsTmp.Filter = "药品ID=" & lng药品ID
                    If rsTmp.EOF Then
                        rsTmp.Filter = 0
                        lng药品ID = Val(rsTmp!药品ID)
                    End If
                    If cbo药房.ListIndex = -1 Then
                        lng药房ID = IIF(mlngPreRow中药房 = 0, mlng中药房, mlngPreRow中药房)
                    Else
                        lng药房ID = cbo药房.ItemData(cbo药房.ListIndex)
                    End If
                    '缺省药房也可能变了（要重新分配规格和数量）
                    Call Get中药房(cbo药房, lng药品ID, mlng病人科室id, mint服务对象, lng药房ID)
                    bln重置药房 = True
                    If cbo药房.ListIndex = -1 And cbo药房.ListCount > 0 Then
                        Call Cbo.SetIndex(cbo药房.hwnd, 0)
                    End If
                End If
            End If
        End If
    End If
    
    '形态变了，要重新分配规格和数量
    Call ReSet中药规格
    If (Not bln重置药房 And mcol规格数量.Count = 1 Or cbo药房.ListIndex = -1) And mint场合 <> 3 Then
        lng药品ID = 0
        strKey = vsExt.Cell(flexcpData, 1, 2)
        If InStr(strKey, "_") > 0 Then
            lng药品ID = Val(Mid(strKey, InStr(strKey, "_") + 1))
        Else
            Set rsTmp = Get中药规格(Val(strKey), Get中药形态)
            If rsTmp.RecordCount > 0 Then
                lng药品ID = Val(rsTmp!药品ID & "")
            End If
        End If
        lng药房ID = IIF(mlngPreRow中药房 = 0, mlng中药房, mlngPreRow中药房)
        Call Get中药房(cbo药房, lng药品ID, mlng病人科室id, mint服务对象, lng药房ID)
        If cbo药房.ListIndex = -1 And cbo药房.ListCount > 0 Then
            Call Cbo.SetIndex(cbo药房.hwnd, 0)
        End If
    End If
End Sub

Private Sub optMode_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        vsExt.SetFocus
    End If
End Sub

Private Sub vsExt_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'功能：非回车确认完后编辑的处理(这里Text:=EditText,但ValidateEdit事件中还没有)
    Dim strPrivs As String, i As Long
    Dim strKey As String, lng药名ID As Long
    
    If Not mblnReturn Then
        If Col Mod 4 = 0 Then '中药
            vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
        ElseIf Col Mod 4 = 1 Then '单味用量
            If Not IsNumeric(vsExt.TextMatrix(Row, Col)) _
                Or Val(vsExt.TextMatrix(Row, Col)) <= 0 _
                Or Val(vsExt.TextMatrix(Row, Col)) > LONG_MAX Then
                vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
            Else
                '小数控制(成套不用)
                If mint场合 <> 3 Then
                    If Val(vsExt.TextMatrix(Row, Col)) <> Int(Val(vsExt.TextMatrix(Row, Col))) Then
                        If mint调用类型 = 1 Then
                            strPrivs = GetInsidePrivs(pm门诊医嘱下达)
                        ElseIf mint调用类型 = 2 Then
                            strPrivs = GetInsidePrivs(pm住院医嘱下达)
                        End If
                        If InStr(strPrivs, "药品小数输入") = 0 Then
                            vsExt.TextMatrix(Row, Col) = IntEx(Val(vsExt.TextMatrix(Row, Col)))
                        End If
                    End If
                End If
                vsExt.Cell(flexcpData, Row, Col) = vsExt.TextMatrix(Row, Col)
                
                strKey = vsExt.Cell(flexcpData, Row, (Col \ 4) * 4 + 2)
                lng药名ID = Val(Split(strKey, "_")(0))
                Call Split中药规格(lng药名ID, Val(vsExt.TextMatrix(Row, Col)), strKey)
                Call Show中药规格(lng药名ID, Val(vsExt.TextMatrix(Row, Col)))
            End If
        ElseIf Col Mod 4 = 3 Then '脚注
            If zlCommFun.ActualLen(vsExt.TextMatrix(Row, Col)) > 100 Then
                vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
            Else
                vsExt.Cell(flexcpData, Row, Col) = vsExt.TextMatrix(Row, Col)
            End If
        End If
    End If
    '确认味数
    Call RefreshWeiNum
End Sub

Private Sub vsExt_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
'功能:显示选择按钮,并保证当前单元格可见
    Dim strKey As String, lng药名ID As Long
    
    If mblnChangeSel = True Then Exit Sub
    '保证当前单元格可见
    If NewRow >= vsExt.FixedRows And NewRow <= vsExt.Rows - 1 Then
        If vsExt.LeftCol >= vsExt.FixedCols And vsExt.LeftCol <= vsExt.Cols - 1 Then
            Call vsExt.ShowCell(NewRow, vsExt.LeftCol)
        End If
    End If

    '显示中药规格
    If Me.Visible Then
        If OldRow <> NewRow Or (OldCol \ 4) <> (NewCol \ 4) Then   '换行或换到另一药品列
            lblZYStock.Caption = ""
            strKey = vsExt.Cell(flexcpData, NewRow, (NewCol \ 4) * 4 + 2)
            If strKey <> "" Then
                lng药名ID = Val(Split(strKey, "_")(0))
                Call Show中药规格(lng药名ID, Val(vsExt.TextMatrix(NewRow, (NewCol \ 4) * 4 + 1)))
            Else
                vs中药规格.Rows = vs中药规格.FixedRows
                cmd形态.Visible = False
            End If
        End If

        If NewCol = (NewCol \ 4) * 4 And vsExt.TextMatrix(NewRow, NewCol) <> "" Then
            strKey = "行【" & vsExt.TextMatrix(NewRow, (NewCol \ 4) * 4) & "】"
        Else
            strKey = "行"
        End If
        vsExt.ToolTipText = "第" & NewRow & strKey
    End If
End Sub

Private Sub vsExt_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, Cancel As Boolean)
   
    If Button = 1 Then
        '单位列鼠标不可进入
        If vsExt.MouseCol Mod 4 = 2 Then Cancel = True
        If mbytUseType = 3 And (vsExt.MouseCol >= 4 Or vsExt.MouseRow >= 2) Then Cancel = True  'mbytUseType = 3时，单药录入
    End If
End Sub

Private Sub vsExt_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)

    If NewCol Mod 4 = 2 Then '单位列按键不可进入
        Cancel = True
        If OldCol > NewCol Then '按键移动时跳过
            vsExt.Col = NewCol - 1
        Else
            vsExt.Col = NewCol + 1
        End If
        vsExt.Row = NewRow
    End If
    
    If mbytUseType = 3 And (NewCol >= 4 Or NewRow >= 2) Then Cancel = True  'mbytUseType = 3时，单药录入
End Sub

Private Sub vsExt_BeforeScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long, Cancel As Boolean)
    If cmd.Visible Then cmd.Visible = False
End Sub

Private Sub vsExt_GotFocus()
    Call vsExt_AfterRowColChange(-1, -1, vsExt.Row, vsExt.Col) '使按钮可见
End Sub

Private Sub vsExt_KeyDown(KeyCode As Integer, Shift As Integer)
'功能：删除数据行
    Dim i As Long, j As Long, k As Long, g As Long
    Dim intRow As Integer        '有效行
    Dim lngRow As Long
    Dim lngCol As Long
    Dim strKey As String, lng药品ID As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim rsTmp As Recordset
    Dim lng药房ID As Long
    
    If KeyCode = vbKeyDelete Then
        If MsgBox("要删除""" & vsExt.TextMatrix(vsExt.Row, (vsExt.Col \ 4) * 4) & """吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        '清除当前味药信息
        
        strKey = vsExt.Cell(flexcpData, vsExt.Row, (vsExt.Col \ 4) * 4 + 2)
        If strKey <> "" Then
            lng药品ID = Val(Split(strKey, "_")(0))
            If InStr(mcol规格数量("_" & strKey), "|") > 0 Then
                vsExt.Select vsExt.Row, (vsExt.Col \ 4) * 4 + 1 '数量列，之前是红色
                vsExt.CellForeColor = vsExt.ForeColorSel
            End If
            mcol规格数量.Remove ("_" & strKey)
            If mint场合 = 3 Then vsExt.Cell(flexcpData, vsExt.Row, (vsExt.Col \ 4) * 4 + 2) = ""
            Call Show中药规格(0, 0)
        End If
        
        For i = 0 To 3
            vsExt.TextMatrix(vsExt.Row, (vsExt.Col \ 4) * 4 + i) = ""
            vsExt.Cell(flexcpData, vsExt.Row, (vsExt.Col \ 4) * 4 + i) = Empty
        Next
        '后面的内容向前移
        For i = vsExt.Row To vsExt.Rows - 1
            For j = 0 To vsExt.Cols - 1 Step 4
                If Not (i = vsExt.Row And j <= (vsExt.Col \ 4) * 4) Then
                    For k = 0 To 3
                        If j = 0 Then
                            vsExt.TextMatrix(i - 1, vsExt.Cols - (4 - k)) = vsExt.TextMatrix(i, j + k)
                            vsExt.Cell(flexcpData, i - 1, vsExt.Cols - (4 - k)) = vsExt.Cell(flexcpData, i, j + k)
                            vsExt.Cell(flexcpForeColor, i - 1, vsExt.Cols - (4 - k)) = vsExt.Cell(flexcpForeColor, i, j + k)
                        Else
                            vsExt.TextMatrix(i, j + k - 4) = vsExt.TextMatrix(i, j + k)
                            vsExt.Cell(flexcpData, i, j + k - 4) = vsExt.Cell(flexcpData, i, j + k)
                            vsExt.Cell(flexcpForeColor, i, j + k - 4) = vsExt.Cell(flexcpForeColor, i, j + k)
                        End If
                        vsExt.TextMatrix(i, j + k) = ""
                        vsExt.Cell(flexcpData, i, j + k) = Empty
                    Next
                End If
            Next
        Next
        '删除多余的空行(至少保留可以最多显示的行数7)
        If vsExt.Rows > 7 Then
            For i = vsExt.Rows - 1 To 7 Step -1
                If Val(vsExt.Cell(flexcpData, i - 1, 2)) = 0 Then
                    vsExt.RemoveItem i
                End If
            Next
        End If
        
        If optMode(1).Enabled = False Then
            If Check相同散装中药 Then
                optMode(1).Enabled = False: optMode(2).Enabled = False
            Else
                optMode(1).Enabled = True: optMode(2).Enabled = True
            End If
        End If
        If vsExt.Row >= vsExt.FixedRows Then Call vsExt_AfterRowColChange(-1, -1, vsExt.Row, vsExt.Col)
        Call vsExt.ShowCell(vsExt.Row, vsExt.Col)
        Call RefreshWeiNum
        '重设药名
        If InStr(strKey, "_") > 0 Then
            Call SetSameItem(Val(Mid(strKey, 1, InStr(strKey, "_") - 1)))
        End If
        
        '如果删除的是第一位中药，重新加载药房(成套编辑不用)
        If mint场合 <> 3 Then
            If cbo药房.ListIndex = -1 And vsExt.Col \ 4 = 0 Then
                lng药品ID = 0
                strKey = vsExt.Cell(flexcpData, 1, 2)
                If InStr(strKey, "_") > 0 Then
                    lng药品ID = Val(Mid(strKey, InStr(strKey, "_") + 1))
                Else
                    Set rsTmp = Get中药规格(Val(strKey), Get中药形态)
                    If rsTmp.RecordCount > 0 Then
                        lng药品ID = Val(rsTmp!药品ID & "")
                    End If
                End If
                lng药房ID = IIF(mlngPreRow中药房 = 0, mlng中药房, mlngPreRow中药房)
                Call Get中药房(cbo药房, lng药品ID, mlng病人科室id, mint服务对象, lng药房ID)
                Call ReSet中药规格(False)
            End If
        End If
    ElseIf KeyCode = vbKeyInsert Then
        If Val(vsExt.Cell(flexcpData, vsExt.Row, (vsExt.Col \ 4) * 4 + 2)) <> 0 Then

            '检查是否有插入的行没有填写
            If CheckIsNullZY(lngRow, lngCol) Then
                MsgBox "请填完整后再插入新项。", vbInformation, Me.Caption
                Call vsExt.Select(lngRow, lngCol)
                Exit Sub
            End If
            '找到有效行
            intRow = -1
            For i = 0 To vsExt.Rows - 1
                For j = 0 To vsExt.Cols - 1 Step 4
                    If vsExt.TextMatrix(i, j) = "" Then
                        intRow = i
                        Exit For
                    End If
                Next
                If intRow <> -1 Then Exit For
            Next
            '如果没有找到有效行列,说明列表已经满了，就添加一行
            If intRow = -1 Then
                intRow = vsExt.Rows - 1
                vsExt.Rows = vsExt.Rows + 1
            End If
            '后面的内容向后移
            For i = intRow To vsExt.Row Step -1
                For j = vsExt.Cols - 1 To 0 Step -4
                    If vsExt.TextMatrix(i, j - 3) <> "" Then
                        If Not (i = vsExt.Row And j <= (vsExt.Col \ 4) * 4) Then
                            For k = 0 To 3
                                If j = vsExt.Cols - 1 Then
                                    vsExt.TextMatrix(i + 1, k) = vsExt.TextMatrix(i, j + (k - 3))
                                    vsExt.Cell(flexcpData, i + 1, k) = vsExt.Cell(flexcpData, i, j + (k - 3))
                                    vsExt.Cell(flexcpForeColor, i + 1, k) = vsExt.Cell(flexcpForeColor, i, j + (k - 3))
                                Else
                                    vsExt.TextMatrix(i, j + k + 1) = vsExt.TextMatrix(i, j + (k - 3))
                                    vsExt.Cell(flexcpData, i, j + k + 1) = vsExt.Cell(flexcpData, i, j + (k - 3))
                                    vsExt.Cell(flexcpForeColor, i, j + k + 1) = vsExt.Cell(flexcpForeColor, i, j + (k - 3))
                                End If
                                vsExt.TextMatrix(i, j + (k - 3)) = ""
                                vsExt.Cell(flexcpData, i, j + (k - 3)) = Empty
                            Next
                        End If
                    End If
                Next
            Next
            For i = 0 To 3
                vsExt.TextMatrix(vsExt.Row, (vsExt.Col \ 4) * 4 + i) = ""
                vsExt.Cell(flexcpData, vsExt.Row, (vsExt.Col \ 4) * 4 + i) = Empty
                vsExt.Cell(flexcpForeColor, vsExt.Row, (vsExt.Col \ 4) * 4 + i) = vsExt.ForeColor
            Next
            Call vsExt.ShowCell(vsExt.Row, vsExt.Col)
            Call Show中药规格(0, 0)
        End If
        Call RefreshWeiNum
    End If
End Sub

Private Sub vsExt_KeyPress(KeyAscii As Integer)
'功能：非编辑状态时，自动移动单元格
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        '定位到下一应输入单元格
        If Val(vsExt.Cell(flexcpData, vsExt.Row, (vsExt.Col \ 4) * 4 + 2)) = 0 Then
            Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        Else
            Call EnterNextCell(vsExt.Row, vsExt.Col)
        End If
    ElseIf KeyAscii = Asc("*") Then
        KeyAscii = 0
        cmd_Click '选择单味中草药或脚注
    End If
End Sub

Private Sub vsExt_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
'功能：输入数据确认
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, str药品 As String
    Dim strStock As String, blnCancel As Boolean, i As Long
    Dim vPoint As PointAPI, strLike As String
    Dim strSamples As String, strPrivs As String
    Dim strKey As String, lng药名ID As Long
    
    If KeyAscii = 13 Then
        mblnReturn = True '标记是按回车确认编辑
        KeyAscii = 0
        
        '优化
        strLike = gstrLike
        If Len(vsExt.EditText) < 2 Then strLike = ""
        
        On Error GoTo errH
        '截取回车后,如果用Msgbox使Edit焦点丢失,则会完成编辑,但不会激活AfterEdit事件
        If Col Mod 4 = 0 Then '中药
            Call Set中药Input(True)
            
            strKey = vsExt.Cell(flexcpData, Row, (Col \ 4) * 4 + 2)
            If strKey <> "" Then lng药名ID = Val(Split(strKey, "_")(0))
            Call Show中药规格(lng药名ID, Val(vsExt.TextMatrix(Row, Col)))
            Exit Sub
        ElseIf Col Mod 4 = 1 Then '单量
            If Not IsNumeric(vsExt.EditText) Or Val(vsExt.EditText) <= 0 Or Val(vsExt.EditText) > LONG_MAX Then
                MsgBox "单味用量输入错误，不是大于零的数字或输入数值过大！", vbInformation, gstrSysName
                vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                Exit Sub
            End If
            '小数控制(成套不用)
            If mint场合 <> 3 Then
                If Val(vsExt.EditText) <> Int(Val(vsExt.EditText)) Then
                    If mint调用类型 = 1 Then
                        strPrivs = GetInsidePrivs(pm门诊医嘱下达)
                    ElseIf mint调用类型 = 2 Then
                        strPrivs = GetInsidePrivs(pm住院医嘱下达)
                    End If
                    If InStr(strPrivs, "药品小数输入") = 0 Then
                        vsExt.EditText = IntEx(Val(vsExt.EditText))
                    End If
                End If
            End If
            vsExt.TextMatrix(Row, Col) = vsExt.EditText
            
            strKey = vsExt.Cell(flexcpData, Row, (Col \ 4) * 4 + 2)
            lng药名ID = Val(Split(strKey, "_")(0))
            
            Call Split中药规格(lng药名ID, Val(vsExt.TextMatrix(Row, Col)), strKey)
            Call Show中药规格(lng药名ID, Val(vsExt.TextMatrix(Row, Col)))
        ElseIf Col Mod 4 = 3 Then '脚注
            If vsExt.EditText <> "" Then
                strSQL = "Select Rownum as ID,编码,名称,简码 From 中药煎服脚注" & _
                    " Where Upper(编码) Like [1] Or Upper(名称) Like [2] Or Upper(简码) Like [2]" & _
                    " Order by 编码"
                vPoint = zlControl.GetCoordPos(vsExt.hwnd, vsExt.CellLeft, vsExt.CellTop)
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "脚注", False, "", "", False, False, True, vPoint.X, vPoint.Y, vsExt.CellHeight, blnCancel, False, True, _
                    UCase(vsExt.EditText) & "%", strLike & UCase(vsExt.EditText) & "%")
            End If
            If rsTmp Is Nothing Then
                If blnCancel Then
                    vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                    Exit Sub
                End If
                '无匹配当作直接输入
                If zlCommFun.ActualLen(vsExt.EditText) > 100 Then
                    MsgBox "脚注输入内容过长，最多只允许 50 个汉字或 100 个字符。", vbInformation, gstrSysName
                    vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                    Exit Sub
                End If
                vsExt.TextMatrix(Row, Col) = vsExt.EditText
            Else
                vsExt.EditText = rsTmp!名称 '直接输入匹配时必要
                vsExt.TextMatrix(Row, Col) = rsTmp!名称
            End If
        End If
        vsExt.Cell(flexcpData, Row, Col) = vsExt.TextMatrix(Row, Col)
        Call EnterNextCell(Row, Col)
    Else
        '单味用量只允许输入数字
        If Col Mod 4 = 1 Then
            If InStr("0123456789." & Chr(8) & Chr(27), Chr(KeyAscii)) = 0 Then
                KeyAscii = 0: Exit Sub
            End If
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsExt_LostFocus()
    If Not ActiveControl Is cmd Then cmd.Visible = False
End Sub

Private Sub vsExt_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsExt.EditSelStart = 0
    vsExt.EditSelLength = zlCommFun.ActualLen(vsExt.EditText)
End Sub

Private Sub vsExt_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'功能：限制某些列不允许编辑(该事件后于BeforeEdit,在EditText赋值之前)
    mblnReturn = False
        
    '必须依次输入
    If Not CellCanEdit(Row, Col) Then Cancel = True
    
    If Col Mod 4 = 1 Then
        vsExt.EditMaxLength = 8
    Else
        vsExt.EditMaxLength = 0
    End If
End Sub

Private Sub vs中药规格_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As ADODB.Recordset
    Dim lng形态 As Long
    Dim strKey As String
    Dim strTmp As String
    Dim i As Long
    Dim str规格数量 As String
    Dim str药品IDs As String
    Dim dbl数量 As Double
    
    With vs中药规格
        If .Col = col规格 Then
            If .ComboData = "" Then
            '没有选择时移开焦点
                .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
            Else
            '散装中药，选择规格之后
                Set rsTmp = .RowData(.FixedRows)
                rsTmp.Filter = "药品ID = " & CLng(.ComboData)
                
                lng形态 = Get中药形态
                If lng形态 = 0 Then
                    strKey = rsTmp!药名ID & "_" & rsTmp!药品ID
                Else
                    strKey = rsTmp!药名ID
                End If
                
                If lng形态 = 0 Then
                    On Error Resume Next
                    strTmp = mcol规格数量("_" & strKey)
                    If err.Number = 0 Then
                        MsgBox "相同规格的药品已存在，请选择其他规格。", vbInformation, gstrSysName
                        .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                        Exit Sub
                    Else
                        err.Clear
                    End If
                    On Error GoTo 0
                Else
                    For i = 1 To .Rows - 1
                        If i <> Row Then
                            If Val(.TextMatrix(i, col药品ID)) = rsTmp!药品ID Then
                                 MsgBox "相同规格的药品已存在，请选择其他规格。", vbInformation, gstrSysName
                                .TextMatrix(Row, Col) = CStr(.Cell(flexcpData, Row, Col))
                                Exit Sub
                            End If
                        End If
                    Next
                End If
                
                .Cell(flexcpData, Row, Col) = CLng(.ComboData)
            
                strTmp = vsExt.Cell(flexcpData, vsExt.Row, (vsExt.Col \ 4) * 4 + 2)
                mcol规格数量.Remove "_" & strTmp
                vsExt.Cell(flexcpData, vsExt.Row, (vsExt.Col \ 4) * 4 + 2) = strKey
                .TextMatrix(Row, col规格) = "" & rsTmp!规格
                .Cell(flexcpData, Row, Col) = "" & rsTmp!规格   '用于恢复
                .TextMatrix(Row, col产地) = "" & rsTmp!产地
                .TextMatrix(Row, col药品ID) = Val(rsTmp!药品ID & "")
                If NVL(rsTmp!是否变价, 0) = 0 Then
                    .TextMatrix(Row, col单价) = Format(CalcPrice(Val("" & rsTmp!药品ID), , , True, , , mstr药品价格等级), gstrDecPrice)
                Else '时价
                    .TextMatrix(Row, col单价) = Format(CalcDrugPrice(Val("" & rsTmp!药品ID), cbo药房.ItemData(cbo药房.ListIndex), Val(.TextMatrix(Row, col数量)), , True, 1, mstr药品价格等级), gstrDecPrice)
                End If
                
                For i = 1 To .Rows - 1
                    str规格数量 = str规格数量 & ";" & .TextMatrix(i, col药品ID) & "," & .TextMatrix(i, col数量)
                    str药品IDs = str药品IDs & "," & .TextMatrix(i, col药品ID)
                    dbl数量 = dbl数量 + Val(.TextMatrix(i, col数量))
                Next
                str规格数量 = Mid(str规格数量, 2)
                str药品IDs = Mid(str药品IDs, 2)
                mcol规格数量.Add str规格数量, "_" & strKey
                If lng形态 = 1 Or lng形态 = 2 Then
                    '饮片和免煎剂修改了规格后根据当前规格重新分配。
                    Call Split中药规格(rsTmp!药名ID, dbl数量, strKey, str药品IDs)
                    vsExt_AfterRowColChange 0, 0, vsExt.Row, vsExt.Col
                End If
                
                '剂量单位不变
                
                '不再设置名称，因为修改规格不会影响名称，只有同一个药品使用了多个规格的，才显示名称
            End If
        End If
    End With
End Sub

Private Sub vs中药规格_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewCol = -1 Or NewRow = -1 Then Exit Sub
    If NewCol = col规格 And vs中药规格.ColComboList(col规格) <> "" Then
        vs中药规格.FocusRect = flexFocusSolid
    Else
        vs中药规格.FocusRect = flexFocusLight
    End If
End Sub

Private Sub vs中药规格_ChangeEdit()
    'Call vs中药规格_AfterEdit(vs中药规格.Row, vs中药规格.Col)
End Sub

Private Sub vs中药规格_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub vs中药规格_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If vs中药规格.ComboIndex <> -1 Then
            Call vs中药规格_KeyPress(13)
        End If
    End If
End Sub

Private Sub vs中药规格_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not (Col = col规格 And vs中药规格.ColComboList(col规格) <> "") Then
        Cancel = True
    End If
End Sub

Private Sub txt付数_GotFocus()
    Call zlControl.TxtSelAll(txt付数)
End Sub

Private Sub txt付数_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf InStr("1234567890" & Chr(8) & Chr(27), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt付数_Validate(Cancel As Boolean)
    '检查输入
    If Not IsNumeric(txt付数.Text) Then
        MsgBox "请输入一个有效的数值。", vbInformation, gstrSysName
        Call zlControl.TxtSelAll(txt付数)
        Cancel = True: Exit Sub
    End If
    If Val(txt付数.Text) <> Int(txt付数.Text) Then
        MsgBox "中药付数应该是整数数值。", vbInformation, gstrSysName
        Call zlControl.TxtSelAll(txt付数)
        Cancel = True: Exit Sub
    End If
    If Val(txt付数.Text) = 0 Then
        MsgBox "请输入一个非零的付数。", vbInformation, gstrSysName
        Call zlControl.TxtSelAll(txt付数)
        Cancel = True: Exit Sub
    End If
    
    If Val(txt付数.Tag) <> Val(txt付数.Text) Then
        txt付数.Tag = Val(txt付数.Text)
    End If
End Sub

Private Function Set中药Input(ByVal blnInputKey As Boolean) As Boolean
'功能：根据输入内容或按*（或按选择按钮）弹出并返回选择的记录集
    Dim rsTmp As ADODB.Recordset, strSQL As String, i As Long
    Dim strMain As String, strPerson As String, strPrivs As String, strCode As String, strLike As String
    Dim str规格 As String, str诊疗 As String, strInput As String, str特性 As String, str存储库房 As String
    Dim lng形态 As Long, lng药房ID As Long, lng药品ID As Long, dbl数量 As Double
    Dim int性别 As Integer, strStock As String
    Dim vPoint As PointAPI, blnCancel As Boolean
    Dim strKey As String
    Dim blnFirst As Boolean '列表中的第一味药
    Dim rs规格 As ADODB.Recordset
    Dim strTsPrivs As String
    
    On Error GoTo errH
    
    If mint场合 <> 3 Then
        If mstr性别 Like "*男*" Then
            int性别 = 1
        ElseIf mstr性别 Like "*女*" Then
            int性别 = 2
        End If
        
        '列表中的第一味药
        If vsExt.Row = vsExt.FixedRows And vsExt.Col = vsExt.FixedCols Then
            If vsExt.TextMatrix(vsExt.FixedRows, vsExt.Col + 4) = "" Then blnFirst = True
        End If
    End If
    
    lng形态 = Get中药形态
    If cbo药房.ListIndex <> -1 Then lng药房ID = cbo药房.ItemData(cbo药房.ListIndex)
    
    '库存,中药房未指定时,读不出库存记录
    If lng药房ID <> 0 Then
        strStock = _
            "Select 药品ID,Sum(Nvl(可用数量,0)) as 库存 From 药品库存" & _
            " Where (Nvl(批次, 0) = 0 Or 效期 Is Null Or 效期 > Trunc(Sysdate))" & _
            " And 性质 = 1 And 库房ID=" & lng药房ID & _
            " Group by 药品ID" & _
            " Having Sum(Nvl(可用数量,0))<>0"
            '中药房一般只有一到两个，此处可不用绑定变量
    Else
        strStock = "Select NULL as 药品ID,NULL as 库存 From Dual"
    End If
        
    If mint场合 <> 3 Then
    '存储库房不能设置病人来源
        If lng形态 = 0 Then
            str存储库房 = " And Exists(select 1 from 收费执行科室 f Where f.收费细目id=d.id and (f.开单科室id Is Null Or f.开单科室id=[8]) And f.执行科室id=" & lng药房ID & ")"
        Else
            str存储库房 = " And Exists(select 1 from 诊疗执行科室 f Where f.诊疗项目id=a.id and (f.开单科室id Is Null Or f.开单科室id=[8]) And f.执行科室id=" & lng药房ID & ")"
        End If
        If blnFirst Then str存储库房 = ""
    End If
    
    '特殊药品权限
    str特性 = ""
    strPrivs = GetInsidePrivs(IIF(mint调用类型 = 1, pm门诊医嘱下达, pm住院医嘱下达))
    strTsPrivs = GetTsPrivs(IIF(mint调用类型 = 1, pm门诊医嘱下达, pm住院医嘱下达))
    
    If mint场合 <> 3 Then
        If InStr(strTsPrivs, "下达麻醉药嘱") = 0 Then
            str特性 = str特性 & " And E.毒理分类<>'麻醉药'"
        End If
        If InStr(strTsPrivs, "下达毒性药嘱") = 0 Then
            str特性 = str特性 & " And E.毒理分类<>'毒性药'"
        End If
        If InStr(strTsPrivs, "下达精神药嘱") = 0 Then
            str特性 = str特性 & " And E.毒理分类 Not IN('精神I类')"
        End If
        If InStr(strTsPrivs, "下达贵重药嘱") = 0 Then
            str特性 = str特性 & " And E.价值分类 Not IN('贵重','昂贵')"
        End If
    End If
    str诊疗 = " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL) And A.服务对象 IN([1],3) And A.类别='7' And Nvl(A.执行频率,0) IN(0,[2]) " & _
            IIF(mint场合 <> 3, " And Nvl(A.适用性别,0) IN(0,[3]) ", "")
        
    If lng形态 = 0 Then
        str规格 = " And Nvl(C.中药形态,0) = [4] And (D.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or D.撤档时间 IS NULL) And D.服务对象 IN([1],3)" & _
                " And (D.站点='" & gstrNodeNo & "' Or D.站点 is Null)"
    Else
         str规格 = " And Exists(Select 1 From 药品规格 F,收费项目目录 S Where F.药品id=C.药品id And F.药品id=S.ID And Nvl(F.中药形态,0) = [4] And S.服务对象 IN([1],3) " & _
                    "And (S.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or S.撤档时间 IS NULL) And (S.站点='" & gstrNodeNo & "' Or S.站点 is Null) )"
    End If
     
    If gblnStock And (mlng中药房 <> 0 And mint场合 <> 3 Or lng药房ID <> 0 And mint场合 = 3) Then
        str规格 = str规格 & " And X.库存>0"
        If blnFirst Then '去掉药房的条件限制
            strStock = "Select 药品ID,Sum(Nvl(可用数量,0)) as 库存 From 药品库存" & _
                " Where (Nvl(批次,0) = 0 Or 效期 Is Null Or 效期 > Trunc(Sysdate))" & _
                " And 性质 = 1 Group by 药品ID Having Sum(Nvl(可用数量,0))<>0"
        End If
    End If
          
    If blnInputKey Then
        strCode = UCase(vsExt.EditText)
        strLike = gstrLike
                
        strInput = " And (A.编码 Like [5] And B.码类=[7]" & _
                    " Or B.名称 Like [6] And B.码类=[7] Or B.简码 Like [6] And B.码类 IN([7],3))"
        If IsNumeric(strCode) Then
            '1X.输入全是数字时只匹配编码'对于药品,则要匹配简码(码类为3的数字码)
            If Mid(gstrMatchMode, 1, 1) = "1" Then strInput = " And (A.编码 Like [5] And B.码类=[7] Or B.简码 Like [6] And B.码类=3)"
        ElseIf zlCommFun.IsCharAlpha(strCode) Then
            'X1.输入全是字母时只匹配简码
            If Mid(gstrMatchMode, 2, 1) = "1" Then strInput = " And B.简码 Like [6] And B.码类=[7]"
        ElseIf zlCommFun.IsCharChinese(strCode) Then
            '包含汉字,则只匹配名称
            strInput = " And B.名称 Like [6] And B.码类=[7]"
        End If
                
        strSQL = "Select Distinct A.ID,A.编码,A.名称,A.计算单位" & _
            " From 诊疗项目目录 A,诊疗项目别名 B" & _
            " Where A.ID=B.诊疗项目ID" & str诊疗 & strInput
               
        strSQL = _
            " Select C.药品ID as ID,D.编码,A.名称,A.计算单位 as 单位,D.规格,D.产地," & _
            IIF(InStr(strPrivs, "显示药品库存") = 0, " Decode(Sign(Nvl(X.库存,0)),1,'有','')", _
                " Decode(X.库存,NULL,NULL,X.库存/" & IIF(mint服务对象 = 1, "C.门诊包装||C.门诊单位)", "C.住院包装||C.住院单位)")) & _
            " as 库存,d.费用类型 As 费用类型,E.处方职务 as 处方职务ID,C.药品ID,A.ID as 药名ID" & _
            " From 药品特性 E,药品规格 C,收费项目目录 D,(" & strSQL & ") A,(" & strStock & ") X" & _
            " Where A.ID=E.药名ID And A.ID=C.药名ID And C.药品ID=D.ID And C.药品ID=X.药品ID(+)" & str特性 & str规格 & str存储库房 & _
            IIF(strLike = "", "", " And Rownum<=100") & _
            " Order by D.编码"
            
        vPoint = zlControl.GetCoordPos(vsExt.hwnd, vsExt.CellLeft, vsExt.CellTop)
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "中药", False, "", "", False, False, True, vPoint.X, vPoint.Y, vsExt.CellHeight, blnCancel, False, True, _
            mint服务对象, IIF(mint期效 = 0, 2, 1), int性别, lng形态, strCode & "%", _
            strLike & strCode & "%", gbytCode + 1, mlng病人科室id)
    Else
        strSQL = "Select 0 as 末级,-1 as ID,-NULL as 上级ID,NULL as 编码," & _
            " CHR(13)||'常用中药' as 名称,NULL as 单位,NULL as 规格,NULL as 产地,NULL as 库存,NULL as 费用类型,NULL as 处方职务ID,NULL as 药品ID,NULL as 药名ID From Dual" & _
            " Union ALL" & _
            " Select 0 as 末级,ID,上级ID,编码,名称,NULL as 单位,NULL as 规格,NULL as 产地,NULL as 库存,NULL as 费用类型,NULL as 处方职务ID,NULL as 药品ID,NULL as 药名ID" & _
            " From 诊疗分类目录 Where 类型=3 And (撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " Start With 上级ID is NULL Connect by Prior ID=上级ID"

         strMain = "Select 1 as 末级,C.药品ID as ID,A.分类ID as 上级ID,D.编码,A.名称,A.计算单位 as 单位,D.规格,D.产地," & _
             IIF(InStr(strPrivs, "显示药品库存") = 0, " Decode(Sign(Nvl(X.库存,0)),1,'有','')", _
                " Decode(X.库存,NULL,NULL,X.库存/" & IIF(mint服务对象 = 1, "C.门诊包装||C.门诊单位)", "C.住院包装||C.住院单位)")) & _
            " as 库存,d.费用类型 As 费用类型,E.处方职务 as 处方职务ID,C.药品ID,A.ID as 药名ID" & _
            " From 诊疗项目目录 A,药品特性 E,药品规格 C,收费项目目录 D,(" & strStock & ") X" & _
            " Where A.ID=E.药名ID And A.ID=C.药名ID And C.药品ID=D.ID And C.药品ID=X.药品ID(+)" & str特性 & str诊疗 & str规格 & str存储库房
            
        strSQL = strSQL & " Union ALL " & strMain
        strPerson = Replace(strMain, "药品特性 E", "药品特性 E,诊疗个人项目 T") & " And T.诊疗项目ID=A.ID And T.人员ID=[5]"
        strSQL = strSQL & " Union ALL " & strPerson
        
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 2, "中药", False, "", "", False, True, False, 0, 0, 0, blnCancel, False, True, _
            mint服务对象, IIF(mint期效 = 0, 2, 1), int性别, lng形态, UserInfo.ID, 0, 0, mlng病人科室id)
    End If
    

    If rsTmp Is Nothing Then
        If Not blnCancel Then
            MsgBox "未找到可用的中药项目，请先到诊疗项目管理中设置。", vbInformation, gstrSysName
        End If
        If blnInputKey Then vsExt.TextMatrix(vsExt.Row, vsExt.Col) = CStr(vsExt.Cell(flexcpData, vsExt.Row, vsExt.Col))
        Exit Function
    End If
                
    '检查重复输入
    If lng形态 = 0 Then
        On Error Resume Next
        strKey = mcol规格数量("_" & rsTmp!药名ID & "_" & rsTmp!药品ID)
        If err.Number = 0 Then
            MsgBox "该味中药在配方中已经录入。", vbInformation, gstrSysName
            If blnInputKey Then vsExt.TextMatrix(vsExt.Row, vsExt.Col) = CStr(vsExt.Cell(flexcpData, vsExt.Row, vsExt.Col))
            Exit Function
        End If
        On Error GoTo 0: err.Clear
    Else
        If ItemExist(rsTmp!药名ID, vsExt.Row, vsExt.Col) Then
            MsgBox "该味中药在配方中已经录入。", vbInformation, gstrSysName
            If blnInputKey Then vsExt.TextMatrix(vsExt.Row, vsExt.Col) = CStr(vsExt.Cell(flexcpData, vsExt.Row, vsExt.Col))
            Exit Function
        End If
    End If
    
    '处方职务检查
    If mint场合 = 0 Then
        strSQL = CheckOneDuty(rsTmp!名称, NVL(rsTmp!处方职务ID), UserInfo.姓名, mbln医保)
        If strSQL <> "" Then
            MsgBox strSQL, vbInformation, gstrSysName
            If blnInputKey Then vsExt.TextMatrix(vsExt.Row, vsExt.Col) = CStr(vsExt.Cell(flexcpData, vsExt.Row, vsExt.Col))
            Exit Function
        End If
    End If
    
    strKey = vsExt.Cell(flexcpData, vsExt.Row, vsExt.Col + 2)
    lng药品ID = -1  '第一味药可能删除了
    If strKey <> "" Then
        If lng形态 = 0 Then  '如果是第一味散装药，规格变了，可用药房跟着改变
            If vsExt.Row = vsExt.FixedRows And vsExt.Col = vsExt.FixedCols Then
                If mcol规格数量("_" & strKey) <> "" Then
                    lng药品ID = Val(Split(mcol规格数量("_" & strKey), ",")(0))
                Else
                    lng药品ID = 0
                End If
            End If
        End If
        mcol规格数量.Remove "_" & strKey
    End If
    
    '获取输入值
    If blnInputKey Then vsExt.EditText = rsTmp!名称 '直接输入匹配时必要
    vsExt.TextMatrix(vsExt.Row, vsExt.Col) = rsTmp!名称
    vsExt.TextMatrix(vsExt.Row, vsExt.Col + 2) = rsTmp!单位
    vsExt.Cell(flexcpData, vsExt.Row, vsExt.Col) = vsExt.TextMatrix(vsExt.Row, vsExt.Col)
    
    If lng形态 = 0 Then
        strKey = rsTmp!药名ID & "_" & rsTmp!药品ID
    Else
        strKey = "" & rsTmp!药名ID '记录中药ID
    End If
    vsExt.Cell(flexcpData, vsExt.Row, vsExt.Col + 2) = strKey
        
    If lng形态 = 0 Then
        '可能在已输入的药品上重输修改，所以不能加条件:If optMode(1).Enabled Then
        If Check相同散装中药 Then
            optMode(1).Enabled = False: optMode(2).Enabled = False
        ElseIf optMode(1).Enabled = False Then
            optMode(1).Enabled = True: optMode(2).Enabled = True
        End If
    End If
    
    If lng形态 = 0 Then
        mcol规格数量.Add rsTmp!药品ID & ",0", "_" & strKey
        
        '散装形态，此时已明确了规格，如果第一味散装药品的规格变了，重设可用药房
        If lng药品ID <> rsTmp!药品ID And vsExt.Row = vsExt.FixedRows And vsExt.Col = vsExt.FixedCols Then
            If cbo药房.ListIndex <> -1 Then
                lng药房ID = cbo药房.ItemData(cbo药房.ListIndex)
            Else
                lng药房ID = IIF(mlngPreRow中药房 = 0, mlng中药房, mlngPreRow中药房)
            End If
            Call Get中药房(cbo药房, Val(rsTmp!药品ID), mlng病人科室id, mint服务对象, lng药房ID)
            If cbo药房.ListIndex = -1 And cbo药房.ListCount > 0 Then Call Cbo.SetIndex(cbo药房.hwnd, 0)
            
            If cbo药房.ListIndex <> -1 Then
                i = cbo药房.ItemData(cbo药房.ListIndex)
            Else
                i = 0
            End If
            If lng药房ID <> i Then Call ReSet中药规格
        End If
    Else
        mcol规格数量.Add "", "_" & rsTmp!药名ID
        If blnFirst Then
            Set rs规格 = Get中药规格(rsTmp!药名ID, lng形态, blnFirst)
            If rs规格.RecordCount > 0 Then
                rs规格.Filter = "中药形态 = " & lng形态
                If rs规格.RecordCount > 0 Then
                    Call Get中药房(cbo药房, Val(rs规格!药品ID), mlng病人科室id, mint服务对象, lng药房ID)
                    If cbo药房.ListIndex = -1 And cbo药房.ListCount > 0 Then Call Cbo.SetIndex(cbo药房.hwnd, 0)
                Else
                    MsgBox "未找到该药品符合要求的形态，请选择其他形态", vbInformation, gstrSysName
                    vsExt.TextMatrix(vsExt.Row, vsExt.Col) = ""
                    Exit Function
                End If
            Else
                MsgBox "未找到该药品任何可用的规格，请选择其他药品", vbInformation, gstrSysName
                vsExt.TextMatrix(vsExt.Row, vsExt.Col) = ""
                Exit Function
            End If
        End If
    End If
    
    '已输入数量时，修改药名
    dbl数量 = Val(vsExt.TextMatrix(vsExt.Row, vsExt.Col + 1))
    If dbl数量 <> 0 Or lng形态 = 0 Then
        Call Split中药规格(Val(rsTmp!药名ID), dbl数量, strKey)
        Call Show中药规格(Val(rsTmp!药名ID), dbl数量)
    End If
    
    '重设名称
    If lng形态 = 0 Then
        Call SetSameItem(Val(rsTmp!药名ID))
    End If
    
    Call EnterNextCell(vsExt.Row, vsExt.Col)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cbo药房_Click()
    '重新分配规格的数量
    If Me.Visible = False Then Exit Sub
    If cbo药房.Tag <> "" And Val(cbo药房.Tag) = cbo药房.ListIndex Then Exit Sub
    cbo药房.Tag = cbo药房.ListIndex
    
    Call ReSet中药规格
End Sub

Private Sub cbo药房_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub cboData_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If cboData.ListIndex <> -1 Then
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    Else
        lngIdx = Cbo.MatchIndex(cboData.hwnd, KeyAscii)
        If lngIdx = -1 And cboData.ListCount > 0 Then lngIdx = 0
        cboData.ListIndex = lngIdx
    End If
End Sub

Private Sub txtJL_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txtJL_GotFocus()
    Call zlControl.TxtSelAll(txtJL)
End Sub

Private Sub cmdInsert_Click()
    Call vsExt_KeyDown(vbKeyInsert, 0)
End Sub

Private Sub cmdOK_Click()
    Dim str中药IDs As String, blnSkip As Boolean
    Dim strMsg As String, strTmp As String
    Dim strSQL As String, i As Long, j As Long
    Dim rsTmp As ADODB.Recordset
    Dim str摘要 As String, str规格数量 As String, lng中药形态 As Long
    Dim strKey As String, lng药名ID As Long
    Dim str手术类型 As String
    Dim blnMsg As Boolean
    
    Dim lngBegin As Long, lngEnd As Long
    Dim strAppend As String, strData As String

    blnSkip = False
    For i = vsExt.FixedRows To vsExt.Rows - 1
        For j = 0 To vsExt.Cols - 1 Step 4
            strKey = vsExt.Cell(flexcpData, i, j + 2)
            If strKey <> "" Then
                str规格数量 = CStr(mcol规格数量("_" & strKey))
                If str规格数量 = "" Then
                    MsgBox "药品""" & vsExt.TextMatrix(i, j) & """未找到可用规格。", vbInformation, gstrSysName
                    vsExt.Select i, j + 1
                    vsExt.SetFocus: Exit Sub
                End If
                If cbo药房.ListIndex = -1 And mint场合 <> 3 Then
                    MsgBox "请选择一个发药药房。", vbInformation, gstrSysName
                    cbo药房.SetFocus: Exit Sub
                End If
                If InStr(str规格数量, "|") > 0 Then
                    MsgBox "数量按规格分配有剩余，请调整""" & vsExt.TextMatrix(i, j) & """的数量或选择用散装规格代替。", vbInformation, gstrSysName
                    vsExt.Select i, j + 1
                    vsExt.SetFocus: Exit Sub
                End If
                If Val(vsExt.TextMatrix(i, j + 1)) = 0 Then
                    If Not blnSkip Then
                        If MsgBox("""" & vsExt.TextMatrix(i, j) & """没有输入单味用量，要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            vsExt.Row = i: vsExt.Col = j + 1
                            Call vsExt.ShowCell(i, j + 1)
                            vsExt.SetFocus: Exit Sub
                        End If
                        blnSkip = True
                    End If
                End If
                If Val(vsExt.TextMatrix(i, j + 1)) <> 0 Then
                    '组合成这种格式：规格ID1,数量,脚注;规格ID3,数量,脚注
                    strTmp = strTmp & ";" & Replace(CStr(mcol规格数量("_" & strKey)), ";", "," & vsExt.TextMatrix(i, j + 3) & ";") & "," & vsExt.TextMatrix(i, j + 3)
                    str中药IDs = str中药IDs & "," & Split(strKey, "_")(0)
                End If
            End If
        Next
    Next
    strTmp = Mid(strTmp, 2)
    str中药IDs = Mid(str中药IDs, 2)
    lng中药形态 = Get中药形态
    
    If strTmp = "" Then
        MsgBox "请在配方中至少输入一味中药。", vbInformation, gstrSysName
        vsExt.Row = vsExt.FixedRows: vsExt.Col = 0
        vsExt.SetFocus: Exit Sub
    End If
    If cboData.ListIndex = -1 Then
        MsgBox "请确定中药配方的煎法。", vbInformation, gstrSysName
        cboData.SetFocus: Exit Sub
    End If
    If cbo药房.ListIndex = -1 And lng中药形态 <> 0 Then
        MsgBox "请确定发药药房。", vbInformation, gstrSysName
        cbo药房.SetFocus: Exit Sub
    End If
    
    '处方职务检查(成套不用）
    If mint场合 = 0 Then
        strSQL = "Select /*+ Rule*/ 药名ID,处方职务 From 药品特性 Where 药名ID IN(Select Column_Value From Table(f_Num2list([1])))"
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str中药IDs)
        For i = vsExt.FixedRows To vsExt.Rows - 1
            For j = 0 To vsExt.Cols - 1 Step 4
                strKey = vsExt.Cell(flexcpData, i, j + 2)
                If strKey <> "" Then
                    lng药名ID = Val(Split(strKey, "_")(0))
                Else
                    lng药名ID = 0
                End If
                
                If lng药名ID <> 0 Then
                    If Val(vsExt.TextMatrix(i, j + 1)) <> 0 Then
                        rsTmp.Filter = "药名ID=" & lng药名ID
                        If Not rsTmp.EOF Then
                            strMsg = CheckOneDuty(vsExt.TextMatrix(i, j), NVL(rsTmp!处方职务), UserInfo.姓名, mbln医保)
                            If strMsg <> "" Then
                                vsExt.Row = i: vsExt.Col = j
                                Call vsExt.ShowCell(i, j)
                                MsgBox strMsg, vbInformation, gstrSysName
                                vsExt.SetFocus: Exit Sub
                            End If
                        End If
                    End If
                End If
            Next
        Next
    End If
    
    '药品禁忌检查（成套不用）
    If mint场合 <> 3 Then
        If Not Check中药禁忌(str中药IDs) Then Exit Sub
    End If
    If cbo药房.ListIndex <> -1 Then
        i = Val(cbo药房.ItemData(cbo药房.ListIndex))
    Else
        i = 0
    End If
    strTmp = strTmp & "|" & cboData.ItemData(cboData.ListIndex) & "|" & lng中药形态 & "|" & Val(txt付数.Text) & "|" & i
    
    '医保信息提示(成套不用）
    If mint场合 <> 3 Then
        If Not mclsInsure Is Nothing And mlng病人ID <> 0 Then  '中药配方
            '医保病人输入内容时的提示
            If UBound(Split(str中药IDs, ",")) = 0 Then
                str摘要 = mclsInsure.GetItemInfo(mint险类, mlng病人ID, Val(Split(mcol规格数量.Item(1), ",")(0)), "", 0, "", str中药IDs & "||" & mint调用类型)
            Else
                str摘要 = mclsInsure.GetItemInfo(mint险类, mlng病人ID, 0, "", 0, "", str中药IDs & "||" & mint调用类型)
            End If
        End If
    End If
    
    strTmp = strTmp & "|" & Trim(txtJL.Text)
    
    If InStr(";" & strTmp, mstr配方明细) > 0 And mstr配方明细 <> "" Then
        strTmp = strTmp & "|1"
    Else
        strTmp = strTmp & "|0"
    End If
    
    mstrExtData = strTmp
    mstr摘要 = str摘要
    mblnOK = True
    Unload Me
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function Check相同散装中药() As Boolean
'功能：检查是否存在相同散装的中药品种
    Dim i As Long, j As Long, strKey As String
    Dim colTmp As New Collection
    
    On Error Resume Next
    With vsExt
        For i = .FixedRows To .Rows - 1
            For j = 0 To .Cols - 1 Step 4
                strKey = .Cell(flexcpData, i, j + 2)
                If strKey <> "" Then
                    colTmp.Add 1, "_" & Split(strKey, "_")(0)
                    If err.Number > 0 Then
                        err.Clear
                        Check相同散装中药 = True
                        Exit Function
                    End If
                End If
            Next
        Next
    End With
End Function

Private Function CheckIsNullZY(ByRef lngRow As Long, ByRef lngCol As Long) As Boolean
'功能：检查是否有已经插入且未填写的
    Dim blnChange As Boolean
    Dim i As Long, j As Long
    
    For i = vsExt.FixedRows To vsExt.Rows - 1
        For j = vsExt.FixedCols To vsExt.Cols - 1 Step 4
            If vsExt.TextMatrix(i, j) = "" Then
                lngRow = i: lngCol = j
                blnChange = True
                Exit For
            End If
        Next
        If blnChange Then Exit For
    Next
    
    If j > vsExt.Cols - 1 Then j = j - 4
    If i > vsExt.Rows - 1 Then i = i - 1
    If j = vsExt.Cols - 4 Then
        If i = vsExt.Rows - 1 Then
            Exit Function
        ElseIf vsExt.TextMatrix(i + 1, vsExt.FixedCols) <> "" Then
            CheckIsNullZY = True
        End If
    Else
        If vsExt.TextMatrix(i, j + 4) <> "" Then
            CheckIsNullZY = True
        End If
    End If
End Function

Private Sub EnterNextCell(ByVal lngRow As Long, ByVal lngCol As Long)
'功能：进入下一个中药配方的输入单元格

    '当前位置未输入中药
    If Val(vsExt.Cell(flexcpData, lngRow, (lngCol \ 4) * 4 + 2)) = 0 Then Exit Sub
    
    '单量未输入
    If lngCol Mod 4 = 1 And vsExt.TextMatrix(lngRow, lngCol) = "" Then Exit Sub
    
    If mbytUseType = 3 And (lngRow > 1 Or lngCol >= 3) Then
        Call zlCommFun.PressKey(vbKeyTab)
        Exit Sub
    End If
            
    If lngCol + 1 <= vsExt.Cols - 1 Then
        lngCol = lngCol + 1
    Else
        If lngRow + 1 > vsExt.Rows - 1 Then
            vsExt.AddItem "", vsExt.Rows
            Call SetSplitLine
        End If
        lngRow = lngRow + 1
        lngCol = vsExt.FixedCols
    End If
    
    vsExt.Row = lngRow: vsExt.Col = lngCol
End Sub

Private Sub cmd_Click()
'功能：打开项目选择器
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim strSQLItem As String, i As Long
    Dim strStock As String, blnCancel As Boolean
    Dim vPoint As PointAPI
    
    On Error GoTo errH
    
    If CellCanEdit(vsExt.Row, vsExt.Col) Then
        If vsExt.Col Mod 4 = 0 Then
            Call Set中药Input(False)
            
        ElseIf vsExt.Col Mod 4 = 3 Then
            '选择脚注
            strSQL = "Select Rownum as ID,编码,名称,简码 From 中药煎服脚注 Order by 编码"
            vPoint = zlControl.GetCoordPos(vsExt.hwnd, vsExt.CellLeft, vsExt.CellTop)
            Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, "脚注", , , , , , True, vPoint.X, vPoint.Y, vsExt.CellHeight, blnCancel, , True)
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "未找到可用的煎服脚注，请先到基础编码管理中设置。", vbInformation, gstrSysName
                End If
                Exit Sub
            End If
            
            '获取输入值
            vsExt.TextMatrix(vsExt.Row, vsExt.Col) = rsTmp!名称
            vsExt.Cell(flexcpData, vsExt.Row, vsExt.Col) = vsExt.TextMatrix(vsExt.Row, vsExt.Col)
            
            Call EnterNextCell(vsExt.Row, vsExt.Col)
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function CellCanEdit(ByVal lngRow As Long, ByVal lngCol As Long) As Boolean
'功能：输入中药配方时,判断指定的单元格当前是否输入内容
'说明：在配方输入表格中,如果前一个未输入,则当前不允许输入
    '定位到上一个中药输入单元
    On Error Resume Next
    
    '如果当前有值，就允许修改
    If Val(vsExt.Cell(flexcpData, lngRow, (lngCol \ 4) * 4)) <> 0 Or vsExt.TextMatrix(lngRow, (lngCol \ 4) * 4) <> "" Then
        CellCanEdit = True
        Exit Function
    End If
    If lngCol = (lngCol \ 4) * 4 + 1 Then '先输药名,再输数量
        If Val(vsExt.Cell(flexcpData, lngRow, (lngCol \ 4) * 4 + 2)) = 0 Then
            CellCanEdit = False
            Exit Function
        End If
    End If
    
    lngCol = (lngCol \ 4) * 4
    If lngCol - 4 >= vsExt.FixedCols Then
        lngCol = lngCol - 4
    Else
        If lngRow - 1 >= vsExt.FixedRows Then
            lngRow = lngRow - 1
            lngCol = vsExt.Cols - 4
        Else
            CellCanEdit = True
            Exit Function
        End If
    End If
    CellCanEdit = Val(vsExt.Cell(flexcpData, lngRow, lngCol + 2)) <> 0
End Function

Private Function ItemExist(ByVal lng中药ID As Long, ByVal lngRow As Long, ByVal lngCol As Long) As Boolean
'功能：判断中药配方输入表格中,指定的中药是否已经输入
    Dim i As Long, j As Long
    
    For i = vsExt.FixedRows To vsExt.Rows - 1
        For j = 0 To vsExt.Cols - 1 Step 4
            If Not (lngRow = i And (lngCol \ 4) * 4 = j) Then
                If Val(vsExt.Cell(flexcpData, i, j + 2)) = lng中药ID Then
                    ItemExist = True
                    Exit Function
                End If
            End If
        Next
    Next
End Function

Private Function Check中药禁忌(ByVal str中药IDs As String) As Boolean
'功能：检查一个配方中的中药配伍禁忌
'参数：str中药IDs="1,2,3,..."
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim str慎用 As String, str禁用 As String, lng组编号 As Long
    
    On Error GoTo errH
    
    strSQL = "Select 组编号 From 诊疗互斥项目" & _
        " Where 项目ID+0 IN(Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist))) Group by 组编号 Having Count(*)>1"
    strSQL = "Select /*+ Rule*/ A.组编号,A.类型,B.名称" & _
        " From 诊疗互斥项目 A,诊疗项目目录 B" & _
        " Where A.项目ID=B.ID And A.组编号 IN(" & strSQL & ")" & _
        " And A.项目ID+0 IN(Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))" & _
        " Order by A.组编号,B.编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str中药IDs)
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            If rsTmp!组编号 <> lng组编号 Then
                If rsTmp!类型 = 1 Then
                    str慎用 = str慎用 & vbCrLf & "●"
                Else
                    str禁用 = str禁用 & vbCrLf & "●"
                End If
                lng组编号 = rsTmp!组编号
            End If
            If rsTmp!类型 = 1 Then
                str慎用 = str慎用 & "，" & rsTmp!名称
            Else
                str禁用 = str禁用 & "，" & rsTmp!名称
            End If
            rsTmp.MoveNext
        Next
        If str禁用 <> "" Then
            MsgBox "当前配方中发现下列药品互相禁用：" & Replace(str禁用, "●，", "● "), vbInformation, gstrSysName
            Exit Function
        ElseIf str慎用 <> "" Then
            If MsgBox("当前配方中发现下列药品互相慎用：" & Replace(str慎用, "●，", "● ") & vbCrLf & vbCrLf & "要继续吗？", _
                vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
    End If
    Check中药禁忌 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub fraBorder_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    If Button = 1 Then
        If Index = 0 Then
            If Me.Height - Y < 2355 Or Me.Height - Y > 7200 Then Exit Sub
            Me.Top = Me.Top + Y
            Me.Height = Me.Height - Y
        ElseIf Index = 1 Then
            If Me.Width + X < 4140 Or Me.Width + X > 9600 Then Exit Sub
            Me.Width = Me.Width + X
        ElseIf Index = 4 Then
            If vsExt.Height + Y < 1000 Or vsExt.Height + Y > Me.Height * 0.7 Then Exit Sub
            vsExt.Height = vsExt.Height + Y
            Call Form_Resize
        End If
    End If
End Sub
