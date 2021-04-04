VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Begin VB.Form frmRegistPlanArrangeNew 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11340
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   20100
   Icon            =   "frmRegistPlanArrangeNew.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11340
   ScaleWidth      =   20100
   StartUpPosition =   2  '屏幕中心
   Begin VB.CheckBox chk立即审核 
      Caption         =   "保存后立即审核"
      Height          =   285
      Left            =   6600
      TabIndex        =   3
      Top             =   120
      Width           =   1650
   End
   Begin VB.PictureBox picBaseBack 
      BorderStyle     =   0  'None
      Height          =   9180
      Left            =   240
      ScaleHeight     =   9180
      ScaleWidth      =   8610
      TabIndex        =   5
      Top             =   480
      Width           =   8610
      Begin VB.PictureBox picBase 
         BorderStyle     =   0  'None
         Height          =   7725
         Left            =   -120
         ScaleHeight     =   7725
         ScaleWidth      =   8130
         TabIndex        =   6
         Top             =   0
         Width           =   8130
         Begin VB.OptionButton opt生效时间 
            Caption         =   "指定时间"
            Height          =   180
            Index           =   1
            Left            =   2040
            TabIndex        =   40
            Top             =   6600
            Value           =   -1  'True
            Width           =   1035
         End
         Begin VB.Frame Frame1 
            Caption         =   "基本信息"
            Height          =   1455
            Left            =   60
            TabIndex        =   27
            Top             =   105
            Width           =   7890
            Begin VB.TextBox txt号别 
               Height          =   300
               IMEMode         =   3  'DISABLE
               Left            =   660
               MaxLength       =   5
               TabIndex        =   34
               Top             =   270
               Width           =   960
            End
            Begin VB.ComboBox cboItem 
               Height          =   300
               IMEMode         =   3  'DISABLE
               Left            =   3900
               Style           =   2  'Dropdown List
               TabIndex        =   33
               Top             =   675
               Width           =   2580
            End
            Begin VB.ComboBox cboDoctor 
               Height          =   300
               Left            =   660
               TabIndex        =   32
               Top             =   1035
               Width           =   2400
            End
            Begin VB.ComboBox cbo科室 
               Height          =   300
               IMEMode         =   3  'DISABLE
               Left            =   660
               Style           =   2  'Dropdown List
               TabIndex        =   31
               Top             =   660
               Width           =   2400
            End
            Begin VB.CheckBox chk病案 
               Caption         =   "挂号时必须建病案"
               Height          =   195
               Left            =   3870
               TabIndex        =   30
               Top             =   1080
               Width           =   1845
            End
            Begin VB.ComboBox cbo号类 
               Height          =   300
               IMEMode         =   3  'DISABLE
               Left            =   3900
               Style           =   2  'Dropdown List
               TabIndex        =   29
               Top             =   270
               Width           =   2595
            End
            Begin VB.CheckBox chk序号控制 
               Caption         =   "序号控制"
               Height          =   255
               Left            =   1750
               TabIndex        =   28
               Top             =   293
               Width           =   1095
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "号别"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Left            =   210
               TabIndex        =   39
               Top             =   330
               Width           =   390
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "科室"
               Height          =   180
               Left            =   240
               TabIndex        =   38
               Top             =   720
               Width           =   360
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "项目"
               Height          =   180
               Left            =   3480
               TabIndex        =   37
               Top             =   750
               Width           =   360
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "医生"
               Height          =   180
               Left            =   240
               TabIndex        =   36
               Top             =   1080
               Width           =   360
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "号类"
               Height          =   180
               Left            =   3465
               TabIndex        =   35
               Top             =   330
               Width           =   360
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "应诊诊室:"
            Height          =   2610
            Left            =   60
            TabIndex        =   21
            Top             =   3840
            Width           =   7875
            Begin VB.OptionButton opt分诊 
               Caption         =   "不分诊"
               Height          =   180
               Index           =   0
               Left            =   1020
               TabIndex        =   25
               Top             =   0
               Value           =   -1  'True
               Width           =   1020
            End
            Begin VB.OptionButton opt分诊 
               Caption         =   "指定诊室"
               Height          =   180
               Index           =   1
               Left            =   2010
               TabIndex        =   24
               Top             =   0
               Width           =   1020
            End
            Begin VB.OptionButton opt分诊 
               Caption         =   "动态分诊"
               Height          =   180
               Index           =   2
               Left            =   3180
               TabIndex        =   23
               Top             =   0
               Width           =   1020
            End
            Begin VB.OptionButton opt分诊 
               Caption         =   "平均分诊"
               Height          =   180
               Index           =   3
               Left            =   4335
               TabIndex        =   22
               Top             =   0
               Width           =   1020
            End
            Begin MSComctlLib.ListView lvwDept 
               Height          =   2190
               Left            =   150
               TabIndex        =   26
               Top             =   300
               Width           =   7605
               _ExtentX        =   13414
               _ExtentY        =   3863
               View            =   2
               Arrange         =   2
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   0   'False
               Checkboxes      =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               Appearance      =   1
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               NumItems        =   0
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "应诊时间"
            Height          =   2070
            Left            =   60
            TabIndex        =   12
            Top             =   1635
            Width           =   7875
            Begin VB.OptionButton opt天 
               Caption         =   "每天(&D)"
               Height          =   315
               Left            =   225
               TabIndex        =   17
               Top             =   285
               Width           =   960
            End
            Begin VB.OptionButton opt周 
               Caption         =   "每周(&W)"
               Height          =   315
               Left            =   225
               TabIndex        =   16
               Top             =   630
               Width           =   930
            End
            Begin VB.ComboBox cbo天 
               Height          =   300
               Left            =   1170
               Style           =   2  'Dropdown List
               TabIndex        =   15
               Top             =   270
               Width           =   1110
            End
            Begin VB.TextBox txt限号 
               Height          =   300
               IMEMode         =   3  'DISABLE
               Left            =   3030
               MaxLength       =   5
               TabIndex        =   14
               Top             =   270
               Width           =   1215
            End
            Begin VB.TextBox txt限约 
               Height          =   300
               IMEMode         =   3  'DISABLE
               Left            =   5145
               MaxLength       =   5
               TabIndex        =   13
               Top             =   270
               Width           =   1215
            End
            Begin VSFlex8Ctl.VSFlexGrid vsPlan 
               Height          =   1275
               Left            =   1170
               TabIndex        =   18
               Top             =   660
               Width           =   6600
               _cx             =   11642
               _cy             =   2249
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
               BackColorBkg    =   -2147483636
               BackColorAlternate=   -2147483643
               GridColor       =   -2147483633
               GridColorFixed  =   -2147483632
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   -2147483643
               FocusRect       =   1
               HighLight       =   1
               AllowSelection  =   -1  'True
               AllowBigSelection=   -1  'True
               AllowUserResizing=   0
               SelectionMode   =   0
               GridLines       =   1
               GridLinesFixed  =   2
               GridLineWidth   =   1
               Rows            =   4
               Cols            =   8
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   300
               RowHeightMax    =   300
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"frmRegistPlanArrangeNew.frx":06EA
               ScrollTrack     =   0   'False
               ScrollBars      =   0
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
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "限号"
               Height          =   180
               Left            =   2595
               TabIndex        =   20
               Top             =   330
               Width           =   360
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "限约"
               Height          =   180
               Left            =   4710
               TabIndex        =   19
               Top             =   330
               Width           =   360
            End
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   0
            Left            =   990
            TabIndex        =   11
            Top             =   7020
            Width           =   2370
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   1
            Left            =   990
            TabIndex        =   10
            Top             =   7410
            Width           =   2370
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   2
            Left            =   5565
            TabIndex        =   9
            Top             =   7020
            Width           =   2370
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   3
            Left            =   5580
            TabIndex        =   8
            Top             =   7410
            Width           =   2370
         End
         Begin VB.OptionButton opt生效时间 
            Caption         =   "立即执行"
            Height          =   360
            Index           =   0
            Left            =   990
            TabIndex        =   7
            Top             =   6525
            Width           =   1530
         End
         Begin MSComCtl2.DTPicker dtpEndDate 
            Height          =   300
            Left            =   5565
            TabIndex        =   41
            Top             =   6555
            Width           =   2385
            _ExtentX        =   4207
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   115736579
            CurrentDate     =   401769
         End
         Begin MSComCtl2.DTPicker dtpBegin 
            Height          =   300
            Left            =   3120
            TabIndex        =   42
            Top             =   6555
            Width           =   2070
            _ExtentX        =   3651
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   115736579
            CurrentDate     =   38091
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "安排人"
            Height          =   180
            Index           =   0
            Left            =   300
            TabIndex        =   48
            Top             =   7080
            Width           =   540
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "安排时间"
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   47
            Top             =   7470
            Width           =   720
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "审核人"
            Height          =   180
            Index           =   2
            Left            =   4950
            TabIndex        =   46
            Top             =   7080
            Width           =   540
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "审核时间"
            Height          =   180
            Index           =   3
            Left            =   4785
            TabIndex        =   45
            Top             =   7470
            Width           =   720
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "计划时间"
            Height          =   180
            Index           =   4
            Left            =   120
            TabIndex        =   44
            Top             =   6600
            Width           =   720
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "～"
            Height          =   180
            Index           =   5
            Left            =   5265
            TabIndex        =   43
            Top             =   6615
            Width           =   180
         End
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   8775
      TabIndex        =   2
      Top             =   1530
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   8775
      TabIndex        =   1
      Top             =   540
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   8775
      TabIndex        =   0
      Top             =   1005
      Width           =   1100
   End
   Begin XtremeSuiteControls.TabControl tbPage 
      Height          =   8700
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   8655
      _Version        =   589884
      _ExtentX        =   15266
      _ExtentY        =   15346
      _StockProps     =   64
   End
End
Attribute VB_Name = "frmRegistPlanArrangeNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit '要求变量声明
'Private mstr计划ID As String, mlng安排ID As Long, mblnSucces As Boolean, mblnFirst As Boolean
'Private mlngModule As Long, mstrPrivs As String
'Private mblnActive As Boolean
'Private Enum mPageIndex
'    EM_计划 = 0
'    EM_时段 = 1
'End Enum
'Private mfrmTime As frmResistPlanTimeSet    '计划时段落
'Private mrsRegOldData As ADODB.Recordset '本地数据集保存,原始挂号安排
'Private mrsRegNewData As ADODB.Recordset '本地数据集保存 重新设置后的安排
'Private mrsRegHistory As ADODB.Recordset '历次挂号的数据集
'Private mblnChangeByCode As Boolean
'Public Enum mRegEditType
'    ed_计划安排 = 0
'    Ed_安排修改 = 1
'    Ed_安排删除 = 2
'    Ed_安排审核 = 3
'    Ed_安排取消 = 4
'    ed_安排查阅 = 5
'End Enum
'Private Enum midxTxt
'    idx_安排人 = 0
'    idx_安排时间 = 1
'    idx_审核人 = 2
'    idx_审核时间 = 3
'End Enum
'Private mEditType As mRegEditType
'Private mstr科室ID As String
'Private mblnCboClick As Boolean     '如果在cbo的keypress事件中用了弹出列表的API函数:sendmessage,当鼠标停在cbo上,输入一个字符,移开焦点或按回车后,
''                                    cbo的值会保存下来,但不会触发click事件,所以需要在validate事件中调用click事件
'Private mrsDoctor As ADODB.Recordset
'
'
'Private Type PlanInfo               '安排改变需要对比的信息
'    str排班         As String       '排班信息
'    str限号         As String       '限号信息
'    bln序号         As Boolean      '是否序号控制
'    bln时间段       As Boolean      '是否设置了时间段
'End Type
'Private mPlanInfo     As PlanInfo '新增时用于保存原始安排信息  修改时 保存原始的计划信息 在保存时 比较相应信息
'Private Enum mPgIndex
'    Pg_计划安排 = 1
'    Pg_计划时段 = 2
'End Enum
'Private Sub InitPage()
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    '功能:初始化页面控件
'    '编制:刘兴洪
'    '日期:2009-09-09 11:01:36
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    Dim i As Long, objItem As TabControlItem, objForm As Object
'    Err = 0: On Error GoTo ErrHand:
'
'    Set objItem = tbPage.InsertItem(mPgIndex.Pg_计划安排, "计划安排", picBaseBack.hWnd, 0)
'    objItem.Tag = mPgIndex.Pg_计划安排
'
'    Set mfrmTime = New frmResistPlanTimeSet
'    Set objItem = tbPage.InsertItem(mPgIndex.Pg_计划时段, "时段设置", mfrmTime.hWnd, 0)
'    objItem.Tag = mPgIndex.Pg_计划时段
'     With tbPage
'        tbPage.Item(0).Selected = True
'        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
'        .PaintManager.BoldSelected = True
'        .PaintManager.Layout = xtpTabLayoutAutoSize
'        .PaintManager.StaticFrame = False
'        .PaintManager.ClientFrame = xtpTabFrameBorder
'    End With
'    Exit Sub
'ErrHand:
'    If ErrCenter = 1 Then
'        Resume
'    End If
'End Sub
'
'
'Public Function ShowCard(ByVal mfrmMain As Form, ByVal lngModule As Long, ByVal strPrivs As String, _
'    ByVal EditType As mRegEditType, Optional lng安排ID As Long, Optional ByVal str计划Id As String = "") As Boolean
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    '功能:显示所要修改的计划安排
'    '入参:mfrmMain-调用的主窗口
'    '     lngModule-模块号
'    '     strPrivs-权限串
'    '     EditType-编辑的类型
'    '     lng安排ID-挂号安排ID.
'    '     str计划Id-安排时为空,否则,否则为指定的计划ID
'    '出参:
'    '返回:成功,返回true,否则返回False
'    '编制:刘兴洪
'    '日期:2009-09-14 14:31:59
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    mEditType = EditType: mlngModule = lngModule: mstrPrivs = strPrivs: mstr计划ID = str计划Id: mblnSucces = False: mlng安排ID = lng安排ID
'    Me.Show 1, mfrmMain
'    ShowCard = mblnSucces
'End Function
'
'Private Function LoadData() As Boolean
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    '功能:加载计划安排数据信息
'    '编制:刘兴洪
'    '日期:2009-09-14 14:40:46
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    Dim rsTemp          As New ADODB.Recordset
'    Dim strSQL          As String
'    Dim i               As Long
'    Dim rs限号          As ADODB.Recordset
'    Dim strTemp         As String
'    Dim bln每周         As Boolean
'    Dim bln限号         As Boolean
'    Dim str限号         As String
'    Dim bln限约         As Boolean
'    Dim str限约         As String
'    Err = 0: On Error GoTo ErrHand:
'
'    '加载安排
'    If mEditType = ed_计划安排 Then
'       '新增安排
'        strSQL = " " & _
'        "   Select A.Id as 安排ID,0 as 计划ID,A.号类,A.项目ID as 计划项目ID,   A.号码,  A.科室id,  A.项目id, A.医生姓名,  A.医生id ,   " & _
'        "          A.周日,  A.周一,  A.周二,  A.周三,  A.周四,  A.周五,  A.周六,A.默认时段间隔, " & _
'        "           A.病案必须,  A.分诊方式,  A.序号控制,  A.开始时间,  A.终止时间,B.名称 As 项目,D.名称 As 科室,NULL　as 生效时间,'3000-01-01 00:00:00' as 失效时间 ," & _
'        "           NULL as 安排人,NULL as 安排时间,NULL 审核人,NULL 审核时间" & _
'        "   From 挂号安排 A,收费项目目录 B,挂号安排计划 C,部门表 D " & _
'        "   Where A.Id=C.安排ID(+) And A.项目id=b.Id(+) And A.科室id =d.Id(+) " & _
'        "         And A.Id=[1]"
'    Else
'         '非新增
'        strSQL = " " & _
'        "Select a.安排id, a.Id As 计划id, a.号类, 计划项目id, a.号码, a.科室id, a.项目id, a.医生姓名, a.医生id,   a.周日, a.周一, a.周二, a.周三," & _
'        "  a.周四, a.周五, a.周六, a.病案必须, a.分诊方式, a.序号控制, a.开始时间, a.终止时间, b.名称 As 项目, d.名称 As 科室, 生效时间, a.失效时间, a.安排人, a.安排时间," & _
'        " a.审核人 , 审核时间,A.默认时段间隔" & _
'        " From (Select c.安排id, c.Id, a.号类, Nvl(c.项目id, a.项目id) As 计划项目id, c.号码, a.科室id, Nvl(c.项目id, a.项目id) As 项目id, C.医生姓名, C.医生id," & _
'        "       c.周日, c.周一, c.周二, c.周三, c.周四, c.周五, c.周六, a.病案必须, c.分诊方式, c.序号控制, a.开始时间, a.终止时间, Nvl(C.默认时段间隔,5) as 默认时段间隔," & _
'        "      To_Char(c.生效时间, 'yyyy-mm-dd hh24:mi:ss') As 生效时间, To_Char(c.失效时间, 'yyyy-mm-dd hh24:mi:ss') As 失效时间, c.安排人," & _
'        "      To_Char(c.安排时间, 'yyyy-mm-dd hh24:mi:ss') As 安排时间, c.审核人, To_Char(c.审核时间, 'yyyy-mm-dd hh24:mi:ss') As 审核时间" & _
'        " From 挂号安排 A, 挂号安排计划 C " & _
'        " Where a.Id = c.安排id) A, 收费项目目录 B, 部门表 D " & _
'        " Where a.项目id = b.Id(+) And a.科室id = d.Id(+) " & _
'        "  and a.id=[2]"
'    End If
'
'    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng安排ID, Val(mstr计划ID))
'    If rsTemp.EOF Then
'        If mEditType = ed_计划安排 Then
'            MsgBox "注意:" & vbCrLf & _
'                   "    挂号安排可能已经被他人删除,不能再进行计划安排", vbInformation + vbOKOnly, gstrSysName
'        Else
'            MsgBox "注意:" & vbCrLf & _
'                   "    挂号计划安排可能已经被他人删除,请检查!", vbInformation + vbOKOnly, gstrSysName
'        End If
'        Exit Function
'    End If
'    If mEditType = ed_计划安排 Then
'        strSQL = "Select 限制项目,限号数,  限约数 From  挂号安排限制 where 安排ID=[1]       "
'    Else
'        strSQL = "Select 限制项目,限号数,  限约数 From  挂号计划限制 where 计划ID=[2]       "
'    End If
'    Set rs限号 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng安排ID, Val(mstr计划ID))
'
'    '检查其他一些功能
'    If mEditType = Ed_安排修改 And Nvl(rsTemp!审核时间) <> "" Then
'            MsgBox "注意:" & vbCrLf & _
'                   "    挂号计划安排已经被他人审核,不能再进行计划修改！", vbInformation + vbOKOnly, gstrSysName
'            Exit Function
'    End If
'    If mEditType = Ed_安排删除 And Nvl(rsTemp!审核时间) <> "" Then
'            MsgBox "注意:" & vbCrLf & _
'                   "    挂号计划安排已经被他人审核,不能再进行计划删除！", vbInformation + vbOKOnly, gstrSysName
'            Exit Function
'    End If
'
'    If mEditType = Ed_安排审核 And Nvl(rsTemp!审核时间) <> "" Then
'            MsgBox "注意:" & vbCrLf & _
'                   "    挂号计划安排已经被他人审核,不能再进行计划审核！", vbInformation + vbOKOnly, gstrSysName
'            Exit Function
'    End If
'
'    If mEditType = Ed_安排取消 And Nvl(rsTemp!审核时间) = "" Then
'            MsgBox "注意:" & vbCrLf & _
'                   "    挂号计划安排已经被他人取消审核,不能再进行计划审核取消！", vbInformation + vbOKOnly, gstrSysName
'            Exit Function
'    End If
'
'    '加载数据到控件中
'    txt号别.Text = Nvl(rsTemp!号码)
'    cbo号类.AddItem Nvl(rsTemp!号类): cbo号类.ListIndex = cbo号类.NewIndex
'    chk序号控制.Value = IIf(Val(Nvl(rsTemp!序号控制)) = 1, 1, 0)
'    '获取的安排或者计划是否序号控制
'    mPlanInfo.bln序号 = IIf(Val(Nvl(rsTemp!序号控制)) = 1, True, False)
'
'    chk病案.Value = IIf(Val(Nvl(rsTemp!病案必须)) = 1, 1, 0)
'
'
'    txtEdit(midxTxt.idx_安排人).Text = Nvl(rsTemp!安排人)
'    txtEdit(midxTxt.idx_安排时间).Text = Nvl(rsTemp!安排时间)
'    If mEditType = ed_计划安排 Then
'        txtEdit(midxTxt.idx_安排人) = UserInfo.姓名
'        txtEdit(midxTxt.idx_安排时间) = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
'    End If
'    txtEdit(midxTxt.idx_审核人) = Nvl(rsTemp!审核人)
'    txtEdit(midxTxt.idx_审核时间) = Nvl(rsTemp!审核时间)
'    If mEditType = Ed_安排审核 Then
'        txtEdit(midxTxt.idx_审核人) = UserInfo.姓名
'        txtEdit(midxTxt.idx_审核时间) = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
'    End If
'
'    With cbo科室
'        .AddItem Nvl(rsTemp!科室): .ItemData(.NewIndex) = Val(Nvl(rsTemp!科室ID)): .ListIndex = .NewIndex
'    End With
'    With cboItem
'         If mEditType = Ed_安排修改 Or mEditType = ed_计划安排 Then
'            zlControl.CboSetText cboItem, rsTemp!项目
'        Else
'            .AddItem Nvl(rsTemp!项目): .ItemData(.NewIndex) = Val(Nvl(rsTemp!项目ID)): .ListIndex = .NewIndex
'        End If
'
'    End With
'    With cboDoctor
'       If mEditType = ed_计划安排 Or mEditType = Ed_安排修改 Then
'          LoadDoctor
'          zlControl.CboSetText cboDoctor, Nvl(rsTemp!医生姓名)
'        Else
'            .AddItem Nvl(rsTemp!医生姓名): .ItemData(.NewIndex) = Val(Nvl(rsTemp!医生ID)): .ListIndex = .NewIndex
'        End If
'    End With
'
'    '加载原始数据到数据集
'     With mrsRegOldData
'        Set mrsRegOldData = New ADODB.Recordset
'        mrsRegOldData.Fields.Append "ID", adBigInt, 18
'        mrsRegOldData.Fields.Append "限制项目", adVarChar, 20
'        mrsRegOldData.Fields.Append "限号数", adBigInt, 10
'        mrsRegOldData.Fields.Append "限约数", adBigInt, 18
'        mrsRegOldData.Fields.Append "序号控制", adBigInt, 18
'        mrsRegOldData.CursorLocation = adUseClient
'        mrsRegOldData.LockType = adLockOptimistic
'        mrsRegOldData.CursorType = adOpenStatic
'        mrsRegOldData.Open
'
'
'        rs限号.Filter = 0
'        If rs限号.RecordCount > 0 Then rs限号.MoveFirst
'        Do While Not rs限号.EOF
'            With mrsRegOldData
'                .AddNew
'                !ID = Val(mstr计划ID)
'                !限制项目 = Nvl(rs限号!限制项目)
'                !限号数 = Val(Nvl(rs限号!限号数))
'                !限约数 = Val(Nvl(rs限号!限约数))
'                !序号控制 = Val(Nvl(rsTemp!序号控制))
'                .Update
'            End With
'            rs限号.MoveNext
'        Loop
'    End With
'
'    Call LoadRegHistory
'    '---------------------------------------------------
'    '判断 每日安排 限号数 限约数 等是否一致
'    '---------------------------------------------------
'    rs限号.Filter = 0
'    If rs限号.RecordCount > 0 Then rs限号.MoveFirst
'
'    bln每周 = Nvl(rsTemp!周日) <> Nvl(rsTemp!周一) Or Nvl(rsTemp!周日) <> Nvl(rsTemp!周二) _
'        Or Nvl(rsTemp!周日) <> Nvl(rsTemp!周三) Or Nvl(rsTemp!周日) <> Nvl(rsTemp!周四) _
'        Or Nvl(rsTemp!周日) <> Nvl(rsTemp!周五) Or Nvl(rsTemp!周日) <> Nvl(rsTemp!周六)
'
'    If bln每周 = False Then
'             rs限号.Filter = "限制项目='周日'"
'             If Not rs限号.EOF Then
'                str限号 = Nvl(rs限号!限号数)
'                str限约 = Nvl(rs限号!限约数)
'             End If
'            For i = 1 To 6
'                strTemp = Switch(i = 0, "日", i = 1, "一", i = 2, "二", i = 3, "三", i = 4, "四", i = 5, "五", True, "六")
'                rs限号.Filter = "限制项目='" & "周" & strTemp & "'"
'                If Not rs限号.EOF Then
'                    bln限号 = Nvl(rs限号!限号数) = str限号
'                    bln限约 = Nvl(rs限号!限约数) = str限约
'                    If bln限约 = False Or bln限号 = False Then Exit For
'                End If
'            Next
'          bln每周 = True
'         If bln限号 And bln限约 Then bln每周 = False
'
'    End If
'
'    If bln每周 Or mrsRegHistory.RecordCount > 0 Then
'        '每周
'        opt周.Value = True:
'        txt限号.Enabled = False: txt限约.Enabled = False
'        With vsPlan
'            For i = 1 To 7
'                '不知什么原因,将.colkey(i)的日,要更改成日日了.
'                strTemp = "周" & Replace(.ColKey(i), "日日", "日")
'                .TextMatrix(1, i) = Nvl(rsTemp.Fields(strTemp))
'                rs限号.Filter = "限制项目='" & strTemp & "'"
'                If Not rs限号.EOF Then
'
'                    .TextMatrix(2, i) = Nvl(rs限号!限号数)
'                    .TextMatrix(3, i) = Nvl(rs限号!限约数)
'                End If
'            Next
'        End With
'    Else
'        '每天
'        opt天.Value = True:  cbo天.ListIndex = GetCboIndex(cbo天, Nvl(rsTemp!周日)): cbo天.Enabled = True
'        If rs限号.RecordCount <> 0 Then rs限号.MoveFirst
'        If rs限号.EOF = False Then
'            txt限号.Text = Nvl(rs限号!限号数)
'            txt限约.Text = Nvl(rs限号!限约数)
'        End If
'    End If
'
'     '------------------------------
'    '获取修改或者新增前的 时间段和 限号数
'    '用于在保存时 对比限号限约、序号控制以及时间段是否发生了变化
'    '如果发生了变化则需要提示  操作员重新设置时段信息
'    '------------------------------
'   mPlanInfo.str排班 = ""
'   mPlanInfo.str限号 = ""
'
'    If bln每周 = False Or mrsRegHistory.RecordCount > 0 Then
'        For i = 1 To 7
'             mPlanInfo.str排班 = mPlanInfo.str排班 & ",'" & Trim(cbo天.Text) & "'"
'             mPlanInfo.str限号 = mPlanInfo.str限号 & "|" & Switch(i = 1, "周日", i = 2, "周一", i = 3, "周二", i = 4, "周三", i = 5, "周四", i = 6, "周五", True, "周六")
'             mPlanInfo.str限号 = mPlanInfo.str限号 & "," & Val(txt限号.Text) & "," & Val(txt限约.Text)
'        Next
'    Else
'        For i = 1 To vsPlan.Cols - 1
'            mPlanInfo.str排班 = mPlanInfo.str排班 & ",'" & Trim(vsPlan.TextMatrix(1, i)) & "'"
'            If Trim(vsPlan.TextMatrix(1, i)) <> "" Then
'                mPlanInfo.str限号 = mPlanInfo.str限号 & "|" & Switch(i = 1, "周日", i = 2, "周一", i = 3, "周二", i = 4, "周三", i = 5, "周四", i = 6, "周五", True, "周六")
'                If Trim(vsPlan.TextMatrix(1, i)) = "" Then
'                     mPlanInfo.str限号 = mPlanInfo.str限号 & ",0,0"
'                Else
'                     mPlanInfo.str限号 = mPlanInfo.str限号 & "," & Val(Trim(vsPlan.TextMatrix(2, i))) & "," & Val(Trim(vsPlan.TextMatrix(3, i)))
'                End If
'            End If
'        Next
'    End If
'    If mPlanInfo.str限号 <> "" Then mPlanInfo.str限号 = Mid(mPlanInfo.str限号, 2)
'    '-------------------------------
'
'    If IsNull(rsTemp!生效时间) Then
'        dtpBegin.Value = Format(zlGetNextWeekDate, "yyyy-mm-dd HH:MM:SS")
'    Else
'        dtpBegin.Value = CDate(Nvl(rsTemp!生效时间))
'    End If
'    dtpEndDate.Value = CDate(Nvl(rsTemp!失效时间, "3000-01-01"))
'
'    Select Case Val(Nvl(rsTemp!分诊方式))     '0-不分诊、1-指定诊室、2-动态分诊、3-平均分诊,对应门诊诊室设置
'        Case 0  '"不分诊"
'            opt分诊(0).Value = True
'        Case 1  ' "指定诊室"
'            opt分诊(1).Value = True
'        Case 2 '"动态分诊"
'            opt分诊(2).Value = True
'        Case 3 ' "平均分诊"
'            opt分诊(3).Value = True
'    End Select
'
'    If mEditType = ed_计划安排 Then
'        strSQL = "Select nvl(生效时间,Sysdate) as 生效时间 ,nvl(失效时间,to_date('3000-01-01','yyyy-mm-dd')) as 失效时间 From 挂号安排计划 where ID=(Select Max(ID) From 挂号安排计划 where 安排ID=[1]) "
'        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng安排ID)
'        If Not rsTemp.EOF Then
'            If Format(rsTemp!失效时间, "yyyy-mm-dd") < "3000-01-01" Then
'                '上一条计划的终止日期,就是本条的生效时间
'                dtpBegin.Value = Format(rsTemp!失效时间, "yyyy-mm-dd HH:MM:SS")
'            Else '以上一条的生效时间的下一周为准
'                dtpBegin.Value = zlGetNextWeekDate(Format(rsTemp!生效时间, "yyyy-mm-dd HH:MM:SS"))
'            End If
'        End If
'
'        strSQL = "Select 号表ID as ID,门诊诊室　From 挂号安排诊室 Where 号表ID=[1]"
'        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng安排ID)
'    Else
'        strSQL = "Select 计划ID as ID,门诊诊室　From 挂号计划诊室 Where 计划ID=[2]"
'        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng安排ID, Val(mstr计划ID))
'    End If
'
'    Do While Not rsTemp.EOF
'        For i = 1 To lvwDept.ListItems.Count
'            If Nvl(rsTemp!门诊诊室) = lvwDept.ListItems(i).Text Then
'                lvwDept.ListItems(i).Checked = True
'            End If
'        Next
'        rsTemp.MoveNext
'    Loop
'    If mEditType = ed_计划安排 Or mEditType = Ed_安排修改 Then mPlanInfo.bln时间段 = Check时段()
'    If mrsRegHistory.RecordCount > 0 Then opt天.Enabled = False
'    LoadData = True
'    Exit Function
'ErrHand:
'    If ErrCenter = 1 Then
'        Resume
'    End If
'    SaveErrLog
'End Function
'Private Function InitData() As Boolean
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    '功能:加载初始化数据
'    '编制:刘兴洪
'    '日期:2009-09-14 15:50:31
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    Dim strSQL As String, rsTemp As New ADODB.Recordset, i As Long
'
'    Err = 0: On Error GoTo ErrHand:
'
'    strSQL = "Select '    ' 时间段 From dual Union All  " & _
'             " Select 时间段 From 时间段"
'
'    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
'    If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
'    Do While Not rsTemp.EOF
'        cbo天.AddItem rsTemp!时间段
'        rsTemp.MoveNext
'    Loop
'
'    With vsPlan
'        .ColComboList(1) = .BuildComboList(rsTemp, "时间段")
'        For i = 2 To .Cols - 1
'            .ColComboList(i) = .ColComboList(1)
'        Next
'        .Tag = .ColComboList(1)
'    End With
'
'
'    '门诊诊室
'    strSQL = "Select 编码,名称　From 门诊诊室 Where (站点='" & gstrNodeNo & "' Or 站点 is Null) Order by 编码"
'    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
'    lvwDept.ListItems.Clear
'    For i = 1 To rsTemp.RecordCount
'        lvwDept.ListItems.Add , "D" & Nvl(rsTemp!编码), Nvl(rsTemp!名称)
'        rsTemp.MoveNext
'    Next
'
'
'    '挂号项目
'    If mEditType = Ed_安排修改 Or mEditType = ed_计划安排 Then
'        strSQL = "Select ID as 序号,名称 From 收费项目目录 " & _
'            " Where 类别='1' And (Sysdate Between 建档时间 And 撤档时间 Or 建档时间<Sysdate And 撤档时间 Is Null)" & _
'            " And (站点='" & gstrNodeNo & "' Or 站点 is Null)" & _
'            " Order by 编码"
'        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
'
'        If rsTemp.EOF Then
'            MsgBox "没有可用的挂号项目信息,请先到挂号项目设置中初始！", vbInformation, gstrSysName
'            Exit Function
'        End If
'
'        cboItem.Clear
'        For i = 1 To rsTemp.RecordCount
'            cboItem.AddItem rsTemp!名称
'            cboItem.ItemData(cboItem.NewIndex) = rsTemp!序号
'            rsTemp.MoveNext
'        Next
'    End If
'
'    'cmdCancel.Caption = "退出(&X)"
'    If mEditType = Ed_安排审核 Then
'        Me.Caption = Me.Caption & "——审核"
'    ElseIf mEditType = Ed_安排删除 Then
'        Me.Caption = Me.Caption & "——删除"
'        'cmdOK.Caption = "删除(&D)"
'    ElseIf mEditType = Ed_安排取消 Then
'        Me.Caption = Me.Caption & "——取消审核"
'    ElseIf mEditType = ed_安排查阅 Then
'        cmdOK.Visible = False
'        cmdCancel.Top = cmdOK.Top
'    End If
'
'    InitData = True
'    Exit Function
'ErrHand:
'    If ErrCenter = 1 Then Resume
'    SaveErrLog
'End Function
'
'
'
'Private Sub Form_Load()
'    Call InitPage
'    opt生效时间(0).Enabled = True: opt生效时间(1).Enabled = True
'    mblnFirst = True
'End Sub
'Private Sub Form_Activate()
'    If mblnFirst = False Then Exit Sub
'    mblnFirst = False
'    If InitData = False Then Unload Me: Exit Sub
'    If LoadData = False Then Unload Me: Exit Sub
'    Call SetCtrlEnabled
'
'    If mEditType = ed_计划安排 Or mEditType = Ed_安排修改 Then
'        zlCtlSetFocus chk序号控制
'    Else
'        zlCtlSetFocus cmdOK
'    End If
'End Sub
'
'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
'End Sub
'Private Sub Form_KeyPress(KeyAscii As Integer)
'    If KeyAscii = Asc("'") Then KeyAscii = 0
'End Sub
'Private Sub SetCtrlEnabled()
'    '设置控件的Enabled属性
'    Dim ctl As Control
'
'    For Each ctl In Me.Controls
'        Select Case UCase(TypeName(ctl))
'        Case "TEXTBOX"
'            ctl.Enabled = False
'            '修改或者新增计划时 开放限号、限约文本框 供修改
'            If ctl Is Me.txt限号 Or ctl Is txt限约 Then
'               ctl.Enabled = mEditType = Ed_安排修改 Or mEditType = ed_计划安排
'            End If
'        Case UCase("ComboBox")
'            If ctl Is cbo天 And mEditType = ed_计划安排 Then
'                   ctl.Enabled = opt天.Value = 1
'              ElseIf ctl Is cboItem Or ctl Is cboDoctor Then
'                 '-----------------------------------------------------
'                 '为修改或者 新增模式时 开放对 项目和医生的更改
'
'                 '------------------------------------------------------
'                   If mEditType = ed_计划安排 Or mEditType = Ed_安排修改 Then
'                       ctl.Enabled = True
'                   Else
'                       ctl.Enabled = False
'                   End If
'               Else:
'                   ctl.Enabled = False
'               End If
'        Case UCase("ListView")
'            ctl.Enabled = False
'        Case UCase("DTPicker")
'            ctl.Enabled = False
'        Case UCase("optionbutton"), UCase("CheckBox")
'            ctl.Enabled = False
'
'        Case Else
'        End Select
'    Next
'
'    Select Case mEditType
'    Case ed_计划安排, Ed_安排修改
'        chk序号控制.Enabled = True
'        txt限号.Enabled = IIf(opt天.Value = True, True, False): txt限约.Enabled = IIf(opt天.Value = True, True, False)
'        cbo天.Enabled = IIf(opt天.Value = True, True, False)
'        dtpBegin.Enabled = IIf(opt生效时间(0).Value = 1, True, False)
'        dtpEndDate.Enabled = True
'        lvwDept.Enabled = True
'        opt分诊(0).Enabled = True: opt分诊(1).Enabled = True: opt分诊(2).Enabled = True: opt分诊(3).Enabled = True
'        opt天.Enabled = True: opt周.Enabled = True
'        dtpBegin.Enabled = True: opt生效时间(0).Enabled = True
'
'        '对分诊进行设置:
'        '   指定医生时，不能设置成,动态分诊或平均分诊
'        If Trim(cboDoctor.Text) <> "" Then
'            opt分诊(2).Enabled = False: opt分诊(3).Enabled = False
'            If opt分诊(2).Value Or opt分诊(3).Value Then opt分诊(0).Value = True
'        Else
'            opt分诊(2).Enabled = True: opt分诊(3).Enabled = True
'        End If
'        If opt天.Value = True Then cbo天.Enabled = True
'    Case Else
'    End Select
'
'    '设置编辑背景色
'    For Each ctl In Me.Controls
'        Select Case UCase(TypeName(ctl))
'        Case "TEXTBOX", UCase("ComboBox")
'            Call zlSetCtrolBackColor(ctl)
'        Case UCase("ListView")
'        Case UCase("DTPicker")
'        Case Else
'        End Select
'    Next
'End Sub
'
'
'Private Sub cmdCancel_Click()
'    Unload Me
'End Sub
'
'Private Sub cmdHelp_Click()
'    ShowHelp App.ProductName, Me.hWnd, Me.Name
'End Sub
'Private Function CheckPlanValied() As Boolean
'    '------------------------------------------------------------------------------------------------------------------------
'    '功能：检查计划的合法性
'    '返回：计划安排合法,返回True,否则返回False
'    '编制：刘兴洪
'    '日期：2010-07-21 17:49:30
'    '说明：
'    '------------------------------------------------------------------------------------------------------------------------
'    Dim rsTemp As New ADODB.Recordset
'    Dim strSQL As String
'    If mEditType <> Ed_安排修改 And mEditType <> ed_计划安排 Then
'        CheckPlanValied = True: Exit Function
'    End If
'
'    If dtpBegin.Value > dtpEndDate.Value Then
'        ShowMsgbox "注意:" & vbCrLf & "    生效时间小于了失效时间,请检查!"
'        If dtpEndDate.Enabled And dtpEndDate.Visible Then dtpEndDate.SetFocus
'        Exit Function
'    End If
'    If zlDatabase.Currentdate > dtpBegin.Value Then
'        ShowMsgbox "注意:" & vbCrLf & "    生效时间小于了当前系统时间,请检查!"
'        If dtpBegin.Enabled And dtpBegin.Visible Then dtpBegin.SetFocus
'        Exit Function
'    End If
'    Set rsTemp = Nothing
'     CheckPlanValied = True: Exit Function
'End Function
'
'Private Function isValied() As Boolean
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    '功能:检查输入的数据的合法性
'    '返回:数据合法,返回true,否则返回False
'    '编制:刘兴洪
'    '日期:2009-09-14 16:31:50
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    Dim rsTemp As ADODB.Recordset, strSQL As String, i As Long, intCount As Integer
'    Dim strTmp As String
'
'    Err = 0: On Error GoTo ErrHand:
'    If Trim(txt号别) = "" Then
'        MsgBox "号别不能为空！", vbInformation, gstrSysName
'        txt号别.SetFocus: Exit Function
'    End If
'    If cbo科室.ListIndex = -1 Then
'        MsgBox "未设置号别所对应的科室！", vbInformation, gstrSysName
'        cbo科室.SetFocus: Exit Function
'    End If
'    If cboItem.ListIndex = -1 Then
'        MsgBox "未设置号别所对应的挂号项目！", vbInformation, gstrSysName
'        cboItem.SetFocus: Exit Function
'    End If
'
'    If opt天.Value Then
'        If cbo天.ListIndex = -1 Then
'            MsgBox "该号别每天的应诊时间未设置！", vbInformation, gstrSysName
'            If txt限号.Enabled Then txt限号.SetFocus
'            Exit Function
'        End If
'        If chk序号控制.Value = 1 Then
'            If Val(txt限号.Text) = 0 And Val(txt限约.Text) = 0 Then
'                MsgBox "使用序号控制时,必须设置限号或限约数！", vbInformation, gstrSysName
'                If txt限号.Enabled Then txt限号.SetFocus
'                Exit Function
'            End If
'        End If
'        '限号限约规则
'        If Trim(txt限号.Text) <> "" Then
'            If Trim(txt限约.Text) <> "" And Val(txt限号.Text) < Val(txt限约.Text) Then
'                MsgBox "限约数应小于限号数！", vbInformation, gstrSysName
'               If txt限约.Enabled Then txt限约.SetFocus
'                Exit Function
'            End If
'        ElseIf Trim(txt限约.Text) <> "" Then
'            MsgBox "限约必须限号！", vbInformation, gstrSysName
'            If txt限号.Enabled Then txt限号.SetFocus
'            Exit Function
'        End If
'    Else
'     With vsPlan
'            strTmp = ""
'            For i = 1 To .Cols - 1
'                If Trim(.TextMatrix(1, i)) <> "" Then
'                    strTmp = strTmp & Trim(vsPlan.TextMatrix(1, i))
'                    If chk序号控制.Value = 1 Then
'                          If Val(.TextMatrix(2, i)) = 0 And Val(.TextMatrix(3, i)) = 0 Then
'                              MsgBox "使用序号控制时,必须设置限号或限约数！", vbInformation, gstrSysName
'                              .Row = 2: .Col = i
'                              .SetFocus: Exit Function
'                          End If
'                      End If
'                        '限号限约规则
'                        If Val(.TextMatrix(2, i)) <> 0 Then
'                            If Trim(.TextMatrix(3, i)) <> "" And Val(.TextMatrix(2, i)) < Val(.TextMatrix(3, i)) Then
'                                MsgBox "限约数应小于限号数！", vbInformation, gstrSysName
'                                .Row = 2: .Col = i
'                                .SetFocus: Exit Function
'                            End If
'                        ElseIf Trim(.TextMatrix(3, i)) <> "" Then
'                            MsgBox "限约必须限号！", vbInformation, gstrSysName
'                            .Row = 2: .Col = i
'                            .SetFocus: Exit Function
'                        End If
'                End If
'            Next
'            If strTmp = "" Then
'                MsgBox "该号别每周的应诊时间未设置！", vbInformation, gstrSysName
'                vsPlan.SetFocus: Exit Function
'            End If
'        End With
'    End If
'    '诊室判断
'    If opt分诊(1).Value Or opt分诊(2).Value Or opt分诊(3).Value Then
'        intCount = 0
'        For i = 1 To lvwDept.ListItems.Count
'            If lvwDept.ListItems(i).Checked Then intCount = intCount + 1
'        Next
'        If opt分诊(1).Value Then
'            If intCount = 0 Then
'                MsgBox "指定诊室时必须选择一个对应的门诊诊室！", vbInformation, gstrSysName
'                lvwDept.SetFocus: Exit Function
'            ElseIf intCount > 1 Then
'                MsgBox "指定诊室时只能选择一个对应的门诊诊室！", vbInformation, gstrSysName
'                lvwDept.SetFocus: Exit Function
'            End If
'        ElseIf opt分诊(2).Value Or opt分诊(3).Value Then
'            If intCount < 2 Then
'                MsgBox "动态分诊或平均分诊时至少要选择两个对应的门诊诊室！", vbInformation, gstrSysName
'                lvwDept.SetFocus: Exit Function
'            End If
'        End If
'    End If
'
'    '项目价格判断
'    If ReadRegistPrice(cboItem.ItemData(cboItem.ListIndex), False, False) = 0 Then
'        MsgBox "项目""" & cboItem.Text & """未设置有效价格,请先到收费项目管理中设置！", vbInformation, gstrSysName
'        cboItem.SetFocus: Exit Function
'    End If
'    If opt生效时间(0).Value = 0 Then
'        If Format(dtpBegin.Value, "yyyy-mm-dd HH:MM:SS") < Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS") Then
'            ShowMsgbox "生效时间不能小于当前系统时间,请检查!"
'            Exit Function
'        End If
'    End If
'    '检查相关的计划
'    If CheckPlanValied = False Then Exit Function
'    Dim blnMulitPlan As Boolean
'    isValied = True
'    Exit Function
'ErrHand:
'    If ErrCenter = 1 Then Resume
'    If 1 = 2 Then
'        Resume
'    End If
'End Function
'
'Private Function SavePlan() As Boolean
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    '功能:保存计划安排
'    '返回:保存成功，返回true,否则返回False
'    '编制:刘兴洪
'    '日期:2009-09-14 16:41:22
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    Dim strSQL As String, str时间段 As String, str诊室 As String, i As Long, int分诊 As Integer
'    Dim lng计划ID As Long, str限号 As String
'    Dim str医生姓名         As String
'    Dim str医生ID           As String
'    Dim blnChange           As Boolean
'    Dim BytType             As Byte
'    Dim vMsgResult          As VbMsgBoxResult
'    Dim strMsg              As String
'    Dim colPro              As Collection
'    'bytType 0-新增时 对时段不进行处理 修改时 对时段只删除已经去掉的排班信息
'    '        1-新增时 提取原安排的时段信息  修改时 对计划的时段进行删除
'
'    Err = 0: On Error GoTo ErrHand:
'
'    str时间段 = "": str限号 = ""
'    If opt天.Value Then
'        For i = 1 To 7
'            str时间段 = str时间段 & ",'" & Trim(cbo天.Text) & "'"
'            str限号 = str限号 & "|" & Switch(i = 1, "周日", i = 2, "周一", i = 3, "周二", i = 4, "周三", i = 5, "周四", i = 6, "周五", True, "周六")
'            str限号 = str限号 & "," & Val(txt限号.Text) & "," & Val(txt限约.Text)
'        Next
'    Else
'        With vsPlan
'            For i = 1 To .Cols - 1
'                str时间段 = str时间段 & ",'" & Trim(.TextMatrix(1, i)) & "'"
'                If Trim(.TextMatrix(1, i)) <> "" Then
'                    str限号 = str限号 & "|" & Switch(i = 1, "周日", i = 2, "周一", i = 3, "周二", i = 4, "周三", i = 5, "周四", i = 6, "周五", True, "周六")
'                    str限号 = str限号 & "," & Val(Trim(vsPlan.TextMatrix(2, i))) & "," & Val(Trim(vsPlan.TextMatrix(3, i)))
'                End If
'            Next
'        End With
'    End If
'    If str限号 <> "" Then str限号 = Mid(str限号, 2)
'
'    If mPlanInfo.bln时间段 Then
'        '判断是已经改变 计划信息
'      blnChange = (mPlanInfo.str排班 <> str时间段) Or (mPlanInfo.str限号 <> str限号) Or (IIf(mPlanInfo.bln序号, 1, 0) <> chk序号控制.Value)
'    End If
'    With lvwDept
'        '取挂号诊室
'        For i = 1 To .ListItems.Count
'            If .ListItems(i).Checked Then
'                str诊室 = str诊室 & ";" & .ListItems(i).Text
'            End If
'        Next
'        str诊室 = Mid(str诊室, 2)
'    End With
'
'    '取分诊方式
'    int分诊 = 0
'    For i = 0 To opt分诊.UBound
'        If opt分诊(i).Value Then int分诊 = i: Exit For
'    Next
'
'    '在计划或者安排设置了时段时 对时段处理的处理类型
''    If mPlanInfo.bln时间段 And mEditType = ed_计划安排 And blnChange = False Then
''        '如果原计划或者安排时 设置了时段 提示操作原进行处理
''        strMsg = "安排中设置了时段,是否提取安排的时段做为计划的时段信息? " & vbCrLf
''        strMsg = strMsg & "[是(Y)]提取安排的时段信息作为计划的时段" & vbCrLf
''        strMsg = strMsg & "[否(N)]不提取安排的时段,重新设置时段" & vbCrLf
''        vMsgResult = MsgBox(strMsg, vbYesNo + vbQuestion + vbDefaultButton1, Me.Caption)
''        BytType = IIf(vMsgResult = vbYes, 1, 0)
''    End If
'    If mEditType = Ed_安排修改 Then
'      BytType = IIf(IIf(mPlanInfo.bln序号, 1, 0) <> chk序号控制.Value, 1, 0)
'    End If
'    '取时间范围
'    If mEditType = ed_计划安排 Then
'        lng计划ID = zlDatabase.GetNextId("挂号安排计划")
'    Else
'        lng计划ID = Val(mstr计划ID)
'    End If
'     If cboDoctor.ListIndex = -1 Then
'        str医生姓名 = ""
'        str医生ID = "0"
'     Else
'        str医生姓名 = cboDoctor.Text
'        str医生ID = Val(cboDoctor.ItemData(cboDoctor.ListIndex))
'     End If
'    'Zl_挂号安排计划_Insert
'    strSQL = "Zl_挂号安排计划_Insert("
'    '  Id_In       In 挂号安排计划.ID%Type,
'    strSQL = strSQL & "" & lng计划ID & ","
'    '  安排id_In   In 挂号安排计划.安排id%Type,
'    strSQL = strSQL & "" & mlng安排ID & ","
'    '  号码_In     In 挂号安排计划.号码%Type,
'    strSQL = strSQL & "'" & txt号别.Text & "',"
'    '  生效时间_In In 挂号安排计划.生效时间%Type,
'    If opt生效时间(0).Value = 1 Then
'        strSQL = strSQL & "to_date('" & Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
'    Else
'        strSQL = strSQL & "to_date('" & dtpBegin.Value & "','yyyy-mm-dd hh24:mi:ss'),"
'    End If
'    '  失效时间_In In 挂号安排计划.失效时间%Type
'    strSQL = strSQL & "to_date('" & dtpEndDate.Value & "','yyyy-mm-dd hh24:mi:ss') "
'    '  周日_In     In 挂号安排计划.周日%Type,
'    '  周一_In     In 挂号安排计划.周一%Type,
'    '  周二_In     In 挂号安排计划.周二%Type,
'    '  周三_In     In 挂号安排计划.周三%Type,
'    '  周四_In     In 挂号安排计划.周四%Type,
'    '  周五_In     In 挂号安排计划.周五%Type,
'    '  周六_In     In 挂号安排计划.周六%Type,
'    strSQL = strSQL & str时间段 & ","
'    '   限号控制_In In Varchar2,
'    strSQL = strSQL & "'" & str限号 & "',"
'    '  分诊方式_In In 挂号安排计划.分诊方式%Type,
'    strSQL = strSQL & "" & int分诊 & ","
'    '  序号控制_In In 挂号安排计划.序号控制%Type,
'    strSQL = strSQL & "" & IIf(chk序号控制.Value = 1, 1, 0) & ","
'    '  项目ID_In   In 挂号安排计划.项目ID%Type,
'    strSQL = strSQL & Me.cboItem.ItemData(cboItem.ListIndex) & ","
'    '医生姓名_In In 挂号安排计划.医生姓名%Type,
'    strSQL = strSQL & IIf(str医生姓名 = "", "NULL,", "'" & str医生姓名 & "',")
'    '医生id_In   In 挂号安排计划.医生id%Type,
'    strSQL = strSQL & str医生ID & ","
'    '  诊室_In     Varchar2,
'    strSQL = strSQL & "'" & str诊室 & "',"
'    '  新增_In Number:=1,处理类型
'    strSQL = strSQL & "" & IIf(mEditType = ed_计划安排, 1, 0) & "," & BytType & ","
'    '立即启用_In Number:=0,
'    strSQL = strSQL & "" & IIf(opt生效时间(0).Value = True, 1, 0) & ")"
'    '立即审核_In Number:=0
'
'    Set colPro = New Collection
'    zlAddArray colPro, strSQL
'    If Not mfrmTime.IsInit Then
'         Call LoadTimePlan
'    End If
'    If mfrmTime.zlSaveData(lng计划ID, colPro) = False Then Exit Function
'    SavePlan = True
'    zlExecuteProcedureArrAy colPro, Me.Caption
'
'    Exit Function
'ErrHand:
'    If ErrCenter = 1 Then Resume
'    SaveErrLog
'End Function
'Private Function SaveVerify() As Boolean
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    '功能:审核挂号安排计划
'    '返回:审核成功,返回true, 否则返回False
'    '编制:刘兴洪
'    '日期:2009-09-14 17:11:24
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    Dim strSQL As String
'    Err = 0: On Error GoTo ErrHand:
'    'Zl_挂号安排计划_Verify(Id_In In 挂号安排计划.ID%Type)
'    strSQL = "Zl_挂号安排计划_Verify(" & Val(mstr计划ID) & ")"
'    zlDatabase.ExecuteProcedure strSQL, Me.Caption
'    SaveVerify = True
'    Exit Function
'ErrHand:
'    If ErrCenter = 1 Then Resume
'    SaveErrLog
'End Function
'Private Function SaveCancel() As Boolean
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    '功能:取消审核挂号安排计划
'    '返回:取消审核成功,返回true, 否则返回False
'    '编制:刘兴洪
'    '日期:2009-09-14 17:11:24
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    Dim strSQL As String
'    Err = 0: On Error GoTo ErrHand:
'    'Zl_挂号安排计划_Cancel(Id_In In 挂号安排计划.ID%Type) Is
'    strSQL = "Zl_挂号安排计划_Cancel(" & Val(mstr计划ID) & ")"
'    zlDatabase.ExecuteProcedure strSQL, Me.Caption
'    SaveCancel = True
'    Exit Function
'ErrHand:
'    If ErrCenter = 1 Then Resume
'     SaveErrLog
'End Function
'Private Function SaveDelete() As Boolean
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    '功能:取消审核挂号安排计划
'    '返回:取消审核成功,返回true, 否则返回False
'    '编制:刘兴洪
'    '日期:2009-09-14 17:11:24
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    Dim strSQL As String
'    Err = 0: On Error GoTo ErrHand:
'    'Zl_挂号安排计划_Delete(Id_In In 挂号安排计划.ID%Type) Is
'    strSQL = "Zl_挂号安排计划_Delete(" & Val(mstr计划ID) & ")"
'    zlDatabase.ExecuteProcedure strSQL, Me.Caption
'    SaveDelete = True
'    Exit Function
'ErrHand:
'    If ErrCenter = 1 Then Resume
'     SaveErrLog
'End Function
'
'Private Sub cmdOK_Click()
'
'
'    If mEditType = ed_安排查阅 Then Unload Me: Exit Sub
'    If mEditType = Ed_安排删除 Then
'        If SaveDelete = False Then Exit Sub
'        mblnSucces = True
'        Unload Me: Exit Sub
'    End If
'
'    If mEditType = Ed_安排审核 Then
'        If SaveVerify = False Then Exit Sub
'        mblnSucces = True
'        Unload Me: Exit Sub
'    End If
'
'    If mEditType = Ed_安排取消 Then
'        If SaveCancel = False Then Exit Sub
'        mblnSucces = True
'        Unload Me: Exit Sub
'    End If
'    If isValied = False Then Exit Sub
'
'    If SavePlan = False Then Exit Sub
'    mblnSucces = True
'    Unload Me
'
'End Sub
'
'Private Sub Form_Resize()
'    Err = 0: On Error Resume Next
'    With cmdOK
'        .Left = ScaleWidth - .Width - 100
'        cmdCancel.Left = .Left
'        cmdHelp.Left = .Left
'    End With
'
'    With tbPage
'        .Top = 50
'        .Height = ScaleHeight - 100
'        .Left = 50
'        .Width = cmdOK.Left - .Left - 100
'    End With
'
'End Sub
'
'Private Sub lvwDept_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'    Dim i As Integer
'    If opt分诊(1).Value Then
'        For i = 1 To lvwDept.ListItems.Count
'            If lvwDept.ListItems(i).Key <> Item.Key Then
'                lvwDept.ListItems(i).Checked = False
'            End If
'        Next
'    End If
'    Set lvwDept.SelectedItem = Item
'End Sub
'Private Sub opt分诊_Click(Index As Integer)
'    Dim i As Integer, strKey As String
'    If opt分诊(1).Value Then
'        For i = 1 To lvwDept.ListItems.Count
'            If lvwDept.ListItems(i).Checked Then
'                If strKey = "" Then
'                    strKey = lvwDept.ListItems(i).Key
'                Else
'                    lvwDept.ListItems(i).Checked = False
'                End If
'            End If
'        Next
'        If strKey <> "" Then
'            Set lvwDept.SelectedItem = lvwDept.ListItems(strKey)
'            lvwDept.SelectedItem.EnsureVisible
'        End If
'    End If
'End Sub
'
'Private Sub opt生效时间_Click(Index As Integer)
'     dtpBegin.Enabled = opt生效时间(0).Value = 0
'End Sub
'
'Private Sub opt天_Click()
'    Dim i As Integer
'    Dim strPlan As String
'    Dim ctl As Control
'
'    With vsPlan
'        For i = 1 To .Cols - 1
'            If Trim(.TextMatrix(1, i)) <> "" Then
'                If strPlan = "" Then
'                    strPlan = .TextMatrix(1, i)
'                Else
'                    If .TextMatrix(1, i) <> strPlan Then
'                        strPlan = "": Exit For
'                    End If
'                End If
'            End If
'        Next
'        For i = 1 To .Cols - 1
'            .TextMatrix(1, i) = ""
'            .TextMatrix(2, i) = ""
'            .TextMatrix(3, i) = ""
'        Next
'        .Enabled = False: .TabStop = False
'    End With
'    opt天.Value = -True: txt限号.Enabled = True: txt限约.Enabled = True
'    cbo天.Enabled = True
'    opt周.Value = False
'    cbo天.ListIndex = GetCboIndex(cbo天, strPlan)
'    cbo天.SetFocus
'
'    '设置编辑背景色
'    For Each ctl In Me.Controls
'        Select Case UCase(TypeName(ctl))
'        Case "TEXTBOX", UCase("ComboBox")
'            Call zlSetCtrolBackColor(ctl)
'        Case UCase("ListView")
'        Case UCase("DTPicker")
'        Case Else
'        End Select
'    Next
'End Sub
'
'Private Sub opt周_Click()
'    Dim i As Integer
'    Dim ctl As Control
'
'    If Trim(cbo天.Text) <> "" Then
'        With vsPlan
'            For i = 0 To .Cols - 1
'                .TextMatrix(1, i) = cbo天.Text
'                .TextMatrix(2, i) = txt限号.Text
'                .TextMatrix(3, i) = txt限约.Text
'            Next
'            .Enabled = True: .TabStop = True
'            .Col = 1: .SetFocus
'        End With
'    End If
'    opt天.Value = False: txt限号.Enabled = False: txt限约.Enabled = False
'    cbo天.Enabled = False: cbo天.ListIndex = -1
'    opt周.Value = True: vsPlan.Enabled = True
'
'    '设置编辑背景色
'    For Each ctl In Me.Controls
'        Select Case UCase(TypeName(ctl))
'        Case "TEXTBOX", UCase("ComboBox")
'            Call zlSetCtrolBackColor(ctl)
'        Case UCase("ListView")
'        Case UCase("DTPicker")
'        Case Else
'        End Select
'    Next
'End Sub
'
'
'
'Private Sub tbPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
'
'
'      If mblnChangeByCode Then Exit Sub
'    PageChange Item
'End Sub
'
'Private Sub PageChange(ByVal Item As XtremeSuiteControls.ITabControlItem)
'
'    If mblnChangeByCode Then Exit Sub
'
'    If Item.Index = mPageIndex.EM_时段 Then
'       mblnChangeByCode = True
'       tbPage.Item(mPageIndex.EM_计划).Selected = True
'        If isValied() = False Then
'            mblnChangeByCode = False
'            Exit Sub
'        End If
'        tbPage.Item(mPageIndex.EM_时段).Selected = True
'        mblnChangeByCode = False
'        Call LoadTimePlan
'    Else
'        If mfrmTime.mblnChange = False Then Exit Sub
'        If mfrmTime.zlPageSelectedChanged() = False Then
'             mblnChangeByCode = True
'            tbPage.Item(mPageIndex.EM_时段).Selected = True
'             mblnChangeByCode = False
'        End If
'    End If
'End Sub
'
'
'
'Private Sub LoadTimePlan()
'    Dim i As Long
'    Dim lng限号数 As Long
'    Dim lng限约数 As Long
'    Dim strTemp As String
'    Dim str安排 As String
'    Dim str排班 As String
'
'    If Not mrsRegNewData Is Nothing Then Set mrsRegNewData = Nothing
'
'    If mrsRegNewData Is Nothing Then
'        Set mrsRegNewData = New ADODB.Recordset
'        mrsRegNewData.Fields.Append "ID", adBigInt, 18
'        mrsRegNewData.Fields.Append "限制项目", adVarChar, 20
'        mrsRegNewData.Fields.Append "排班", adVarChar, 20
'        mrsRegNewData.Fields.Append "限号数", adBigInt, 10
'        mrsRegNewData.Fields.Append "限约数", adBigInt, 18
'        mrsRegNewData.Fields.Append "序号控制", adBigInt, 18
'        mrsRegNewData.CursorLocation = adUseClient
'        mrsRegNewData.LockType = adLockOptimistic
'        mrsRegNewData.CursorType = adOpenStatic
'        mrsRegNewData.Open
'     End If
'
'     If opt天.Value = True Then
'          lng限号数 = Val(txt限号.Text)
'          lng限约数 = Val(txt限约.Text)
'          str排班 = Me.cbo天.Text
'          For i = 0 To 6
'            strTemp = Switch(i = 0, "周日", i = 1, "周一", i = 2, "周二", i = 3, "周三", i = 4, "周四", i = 5, "周五", i = 6, "周六")
'            '周一,限号数,限约数|周二,限号数,限约数|....
'            str安排 = str安排 & "|" & strTemp & "," & lng限号数 & "," & lng限约数
'             With mrsRegNewData
'                .AddNew
'                !ID = Val(mstr计划ID)
'                !限制项目 = strTemp
'                !排班 = str排班
'                !限号数 = lng限号数
'                !限约数 = lng限约数
'                !序号控制 = Me.chk序号控制.Value
'                .Update
'            End With
'          Next
'
'        Else
'
'           With vsPlan
'            For i = 1 To .Cols - 1
'                If Trim(.TextMatrix(1, i)) <> "" Then
'                    strTemp = Switch(i = 1, "周日", i = 2, "周一", i = 3, "周二", i = 4, "周三", i = 5, "周四", i = 6, "周五", True, "周六")
'                    lng限号数 = Val(Trim(vsPlan.TextMatrix(2, i)))
'                    lng限约数 = Val(Trim(vsPlan.TextMatrix(3, i)))
'                    str排班 = Trim(vsPlan.TextMatrix(1, i))
'                    str安排 = str安排 & "|" & strTemp & "," & lng限号数 & "," & lng限约数
'                    With mrsRegNewData
'                        .AddNew
'                        !ID = Val(mstr计划ID)
'                        !限制项目 = strTemp
'                        !排班 = str排班
'                        !限号数 = lng限号数
'                        !限约数 = lng限约数
'                        !序号控制 = Me.chk序号控制.Value
'                        .Update
'                    End With
'                End If
'            Next
'        End With
'     End If
'     If str安排 <> "" Then str安排 = Mid(str安排, 2)
''Public Enum mRegEditType
''Ed_计划安排 = 0
''Ed_安排修改 = 1
''Ed_安排删除 = 2
''Ed_安排审核 = 3
''Ed_安排取消 = 4
''Ed_安排查阅 = 5
''End Enum
'
'     mfrmTime.zlShowPagePlan str安排, mrsRegNewData, mrsRegHistory, chk序号控制.Value = 1, Switch(mEditType = ed_计划安排, EM_计划_增加, mEditType = Ed_安排修改, EM_计划_修改, True, EM_计划_查阅), mlng安排ID, Val(mstr计划ID)
'End Sub
'
''Private Sub tbPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
''    Dim i As Long
''    Dim lng限号数 As Long
''    Dim lng限约数 As Long
''    Dim strTemp As String
''    Dim str安排 As String
''    If Item.Index <> mPageIndex.EM_时段 Then Exit Sub
''    If Not mrsRegNewData Is Nothing Then Set mrsRegNewData = Nothing
''    If mrsRegNewData Is Nothing Then
''        With mrsRegNewData
''        Set mrsRegNewData = New ADODB.Recordset
''        mrsRegNewData.Fields.Append "ID", adBigInt, 18
''        mrsRegNewData.Fields.Append "限制项目", adVarChar, 20
''        mrsRegNewData.Fields.Append "限号数", adBigInt, 10
''        mrsRegNewData.Fields.Append "限约数", adBigInt, 18
''        mrsRegNewData.Fields.Append "序号控制", adBigInt, 18
''        mrsRegNewData.CursorLocation = adUseClient
''        mrsRegNewData.LockType = adLockOptimistic
''        mrsRegNewData.CursorType = adOpenStatic
''        mrsRegNewData.Open
''        If opt天.Value = True Then
''          lng限号数 = Val(txt限号.Text)
''          lng限约数 = Val(txt限约.Text)
''          For i = 0 To 6
''            strTemp = Switch(i = 0, "周日", i = 1, "周一", i = 2, "周二", i = 3, "周三", i = 4, "周四", i = 5, "周五", i = 6, "周六")
''            .AddNew
''            !ID = Val(mstr计划ID)
''            !限制项目 = strTemp
''            !限号数 = lng限号数
''            !限约数 = lng限约数
''            !序号控制 = Me.chk序号控制.Value
''            .Update
''          Next
''
''        Else
''
''           With vsPlan
''            For i = 1 To .Cols - 1
''                If Trim(.TextMatrix(1, i)) <> "" Then
''                    strTemp = Switch(i = 1, "周日", i = 2, "周一", i = 3, "周二", i = 4, "周三", i = 5, "周四", i = 6, "周五", True, "周六")
''                    lng限号数 = Val(Trim(vsPlan.TextMatrix(2, i)))
''                    lng限约数 = Val(Trim(vsPlan.TextMatrix(3, i)))
''                    With mrsRegNewData
''                        .AddNew
''                        !ID = Val(mstr计划ID)
''                        !限制项目 = strTemp
''                        !限号数 = lng限号数
''                        !限约数 = lng限约数
''                        !序号控制 = Me.chk序号控制.Value
''                        .Update
''                    End With
''                End If
''            Next
''        End With
''
''        End If
''    End With
''
''    End If
''    If mfrmTime Is Nothing Then
''        Set mfrmTime = New frmResistPlanTimeSet
''    End If
''End Sub
'
'Private Sub txt号别_GotFocus()
'    zlControl.TxtSelAll txt号别
'End Sub
'Private Sub txt号别_KeyPress(KeyAscii As Integer)
'    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
'End Sub
'
'Private Sub txt限号_GotFocus()
'    zlControl.TxtSelAll txt限号
'End Sub
'Private Sub txt限号_KeyPress(KeyAscii As Integer)
'    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
'End Sub
'
'Private Sub txt限号_Validate(Cancel As Boolean)
'    If Trim(txt限号.Text) = "" And Trim(txt限约.Text) <> "" Then
'        MsgBox "限约必须限号!", vbInformation, gstrSysName
'        Cancel = True: Exit Sub
'    End If
'    If Trim(txt限号.Text) <> "" And Trim(txt限约.Text) <> "" And Val(txt限号.Text) < Val(txt限约.Text) Then
'        MsgBox "限约数不能小于限号数!", vbInformation, gstrSysName
'        Cancel = True: Exit Sub
'    End If
'End Sub
'Private Sub txt限约_GotFocus()
'    zlControl.TxtSelAll txt限约
'End Sub
'
'Private Sub txt限约_KeyPress(KeyAscii As Integer)
'    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
'    If Val(txt限号.Text) = 0 Then KeyAscii = 0
'End Sub
'
'Private Sub txt限约_Validate(Cancel As Boolean)
'    If Val(txt限号.Text) < Val(txt限约.Text) And _
'        Trim(txt限号.Text) <> "" And Trim(txt限约.Text) <> "" Then
'        MsgBox "限约数不能小于限号数!", vbInformation, gstrSysName
'        Cancel = True: Exit Sub
'    End If
'End Sub
'
'Private Function zlCheckRegistPlanIsValied(ByRef blnMulitNumPlan As Boolean) As Boolean
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    '功能:检查当前所输入的号码是否合法
'    '出参:blnMulitNumPlan-返回是否有多个相同(同一项目,同一科室,同一人,不同号)的安排
'    '返回:合法返回,则返回true,否则返回False
'    '编制:刘兴洪
'    '日期:2010-12-29 10:26:45
'    '检查规则（同一项目,同一科室,同一人,不同号）:
'    '     1.同天内不能有交叉的安排
'    '问题目:35057
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    Dim strSQL As String, rsTemp As ADODB.Recordset, str医生 As String
'    Dim lng项目id As Long, lng科室ID As Long, lng医生ID As Long
'    Dim str号别 As String, strTemp As String, strTemp1 As String
'    Dim i As Long, bytCheckType As Byte '0-检查计划是否合法;1-检查安排中正在执行项目是否合法.
'    Dim strTittle As String
'
'    On Error GoTo errHandle
'    lng科室ID = cbo科室.ItemData(cbo科室.ListIndex)
'    lng项目id = cboItem.ItemData(cboItem.ListIndex)
'    lng医生ID = 0: str医生 = Trim(cboDoctor.Text)
'    If cboDoctor.ListIndex <> -1 Then lng医生ID = cboDoctor.ItemData(cboDoctor.ListIndex)
'
'    '检查计划中是否存在重复
'    bytCheckType = 0
'goReCheck:
'    If bytCheckType <> 0 Then
'
'        strSQL = "" & _
'        "   Select Distinct A.号码, A.周日 D0, A.周一 D1, A.周二 D2, A.周三 D3, A.周四 D4, A.周五 D5, A.周六 D6, " & _
'        "                 Nvl(To_Char(a.开始时间, 'YYYY-MM-DD HH24:MI:SS'), '1901-01-01') 生效时间, " & _
'        "                 Nvl(To_Char(a.终止时间, 'YYYY-MM-DD HH24:MI:SS'), '3000-01-01 00:00:00') 失效时间 " & _
'        "   From 挂号安排 A,挂号安排 B " & _
'        "   Where A.科室id = b.科室id And A.医生姓名 = b.医生姓名 And Nvl(A.医生id, 0) = nvl(b.医生id,0) " & _
'        "               And a.ID + 0 <> [1]   And B.ID = [1]  " & _
'        "   Order By 号码"
'            strTittle = "安排"
'    Else
'        strSQL = "" & _
'            "   Select  distinct A.号码,A.周日 D0,A.周一 D1,A.周二 D2,A.周三 D3,A.周四 D4,A.周五 D5,A.周六 D6," & _
'            "           To_Char(A.生效时间,'YYYY-MM-DD HH24:MI:SS') 生效时间,To_Char(A.失效时间,'YYYY-MM-DD HH24:MI:SS') 失效时间" & _
'            "   From 挂号安排计划 A, 挂号安排 B,挂号安排 C " & _
'            "   Where A.安排ID=B.ID and B.科室ID=C.科室ID and B.医生姓名=C.医生姓名 and nvl(B.医生ID,0)=nvl(C.医生ID,0) " & _
'            "           And B.ID+0<>[1] and C.ID=[1]  " & _
'            "   Order by 号码"
'            strTittle = "计划安排"
'    End If
'    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng安排ID)
'    blnMulitNumPlan = Not rsTemp.EOF
'    If blnMulitNumPlan = False And bytCheckType = 0 Then
'        bytCheckType = bytCheckType + 1
'        GoTo goReCheck:
'    End If
'    If blnMulitNumPlan = False Then zlCheckRegistPlanIsValied = True: Exit Function
'    str号别 = ""
'    Do While Not rsTemp.EOF
'        str号别 = str号别 & "," & Nvl(rsTemp!号码)
'        If (Nvl(rsTemp!生效时间) >= Format(dtpBegin.Value, "yyyy-mm-dd HH:MM:SS") And Nvl(rsTemp!生效时间) < Format(dtpEndDate.Value, "yyyy-mm-dd HH:MM:SS")) Or _
'           (Nvl(rsTemp!失效时间) >= Format(dtpBegin.Value, "yyyy-mm-dd HH:MM:SS") And Nvl(rsTemp!失效时间) < Format(dtpEndDate.Value, "yyyy-mm-dd HH:MM:SS")) Or _
'           (Format(dtpBegin.Value, "yyyy-mm-dd HH:MM:SS") >= Nvl(rsTemp!生效时间) And Format(dtpBegin.Value, "yyyy-mm-dd HH:MM:SS") < Nvl(rsTemp!失效时间)) Or _
'           (Format(dtpEndDate.Value, "yyyy-mm-dd HH:MM:SS") >= Nvl(rsTemp!生效时间) And Format(dtpEndDate.Value, "yyyy-mm-dd HH:MM:SS") < Nvl(rsTemp!失效时间)) Then
'           '时间内不能交叉
'            If opt天.Value Then
'                If Trim(Nvl(rsTemp!D0)) <> "" Then strTemp = strTemp & vbCrLf & "  周日:" & Nvl(rsTemp!D0)
'                If Trim(Nvl(rsTemp!D1)) <> "" Then strTemp = strTemp & vbCrLf & "  周一:" & Nvl(rsTemp!D1)
'                If Trim(Nvl(rsTemp!D2)) <> "" Then strTemp = strTemp & vbCrLf & "  周二:" & Nvl(rsTemp!D2)
'                If Trim(Nvl(rsTemp!D3)) <> "" Then strTemp = strTemp & vbCrLf & "  周三:" & Nvl(rsTemp!D3)
'                If Trim(Nvl(rsTemp!D4)) <> "" Then strTemp = strTemp & vbCrLf & "  周四:" & Nvl(rsTemp!D4)
'                If Trim(Nvl(rsTemp!D5)) <> "" Then strTemp = strTemp & vbCrLf & "  周五:" & Nvl(rsTemp!D5)
'                If Trim(Nvl(rsTemp!D6)) <> "" Then strTemp = strTemp & vbCrLf & "  周六:" & Nvl(rsTemp!D6)
'                If strTemp <> "" Then
'                    strTemp = vbCrLf & "在号别 [" & rsTemp!号码 & "] 中已有如下" & strTittle & ":" & vbCrLf & "        " & Mid(strTemp, 2) & vbCrLf & vbCrLf & "  生效时间:" & IIf(Nvl(rsTemp!生效时间) = "1901-01-01", "无限", Nvl(rsTemp!生效时间) & "-" & Nvl(rsTemp!失效时间)) & vbCrLf
'                    Call MsgBox("发现『" & cboDoctor.Text & "』医生存在与当前号别重复或交叉的挂号计划安排 " & vbCrLf & strTemp & vbCrLf & vbCrLf & "请修改此计划安排.", vbInformation + vbOKOnly + vbDefaultButton2, gstrSysName)
'                    zlCheckRegistPlanIsValied = False: Exit Function
'                End If
'            Else
'                With vsPlan
'                    For i = 0 To 6
'                        strTemp1 = "  周" & Switch(i = 0, "日", i = 1, "一", i = 2, "二", i = 3, "三", i = 4, "四", i = 5, "五", True, "六")
'                        If Trim(Nvl(rsTemp.Fields("D" & i).Value)) <> "" And Trim(.TextMatrix(1, i)) <> "" Then
'                            '存在,肯定重复了
'                            strTemp = strTemp & vbCrLf & strTemp1 & ":" & Trim(Nvl(rsTemp.Fields("D" & i).Value))
'                        End If
'                    Next
'                End With
'                If strTemp <> "" Then
'                    strTemp = vbCrLf & "在号别 [" & rsTemp!号码 & "] 中已有如下" & strTittle & ":" & vbCrLf & "        " & Mid(strTemp, 2) & vbCrLf & "  生效时间:" & IIf(Nvl(rsTemp!生效时间) = "1901-01-01", "无限", Nvl(rsTemp!生效时间) & "-" & Nvl(rsTemp!失效时间)) & vbCrLf
'                    Call MsgBox("发现『" & cboDoctor.Text & "』医生存在与当前号别重复或交叉的挂号安排 " & vbCrLf & strTemp & vbCrLf & vbCrLf & "请修改此计划安排.", vbInformation + vbOKOnly + vbDefaultButton2, gstrSysName)
'                    zlCheckRegistPlanIsValied = False: Exit Function
'                End If
'            End If
'        End If
'        rsTemp.MoveNext
'    Loop
'    If bytCheckType = 0 Then
'        bytCheckType = bytCheckType + 1
'        GoTo goReCheck:
'    End If
'    zlCheckRegistPlanIsValied = True
'    Exit Function
'errHandle:
'    If ErrCenter() = 1 Then
'        Resume
'    End If
'     SaveErrLog
'End Function
'
'Private Sub vsPlan_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'    With vsPlan
'        If mEditType <> ed_计划安排 And mEditType <> Ed_安排修改 Then Cancel = True: Exit Sub
'        If Not opt周.Value = True Then Cancel = True: Exit Sub
'    End With
'End Sub
'Private Sub vsPlan_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'      '---------------------------------------------------------------------------------------------------------------------------------------------
'    '功能:设置相关的格式
'    '编制:刘兴洪
'    '日期:2011-11-11 11:33:11
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    With vsPlan
'        If Row = 1 Then
'              If Trim(.EditText) = "" Then
'               .TextMatrix(2, Col) = ""
'               .TextMatrix(3, Col) = ""
'            End If
'            Exit Sub
'        End If
'        .TextMatrix(Row, Col) = Format(Val(.TextMatrix(Row, Col)), "###;;;")
'    End With
'    Exit Sub
'End Sub
'Private Sub vsPlan_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
'   Call zl_VsGridRowChange(vsPlan, OldRow, NewRow, OldCol, NewCol)
'    vsPlan.ColComboList(NewCol) = ""
'    If OldRow = 1 And Trim(vsPlan.TextMatrix(1, OldCol)) = "" Then
'        vsPlan.TextMatrix(2, OldCol) = ""
'        vsPlan.TextMatrix(3, OldCol) = ""
'    End If
'    If OldRow = 2 And Trim(vsPlan.TextMatrix(3, OldCol)) = "" Then
'        vsPlan.TextMatrix(3, OldCol) = vsPlan.TextMatrix(2, OldCol)
'    End If
'    If NewRow <> 1 Then Exit Sub
'    vsPlan.ColComboList(NewCol) = vsPlan.Tag
'End Sub
'Private Sub vsPlan_GotFocus()
'    Call zl_VsGridGotFocus(vsPlan)
'End Sub
'Private Sub vsPlan_KeyDown(KeyCode As Integer, Shift As Integer)
'    Dim lngCol As Long, blnCancel As Boolean, lngRow As Long
'    With vsPlan
'        If KeyCode = vbKeyDelete Then
'            .TextMatrix(.Row, .Col) = ""
'        End If
'    End With
'    If KeyCode <> vbKeyReturn Then Exit Sub
'
'    With vsPlan
'        If .Row = 3 And .Col = .Cols - 1 Then zlCommFun.PressKey vbKeyTab: Exit Sub
'        If .Row < 3 Then
'            .Row = .Row + 1
'        Else
'            .Row = 1
'            If .Col + 1 <= .Cols - 1 Then .Col = .Col + 1
'         End If
'    End With
'End Sub
'
'Private Sub vsPlan_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
'    '编辑处理
'    Dim intCol As Integer, strKey As String, lngRow As Long
'
'    If KeyCode <> vbKeyReturn Then Exit Sub
'    With vsPlan
'            If .Row = 3 And .Col = .Cols - 1 Then zlCommFun.PressKey vbKeyTab: Exit Sub
'        If .Row < 3 Then
'            .Row = .Row + 1
'        Else
'            .Row = 1
'            If .Col + 1 <= .Cols - 1 Then .Col = .Col + 1
'         End If
'    End With
'End Sub
'Private Sub vsPlan_KeyPress(KeyAscii As Integer)
'    If KeyAscii = vbKeyReturn Then KeyAscii = 0
'End Sub
'Private Sub vsPlan_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
'    With vsPlan
'        If Row <= 1 Then Exit Sub
'        VsFlxGridCheckKeyPress vsPlan, Row, Col, KeyAscii, m数字式
'    End With
'End Sub
'Private Sub vsPlan_LostFocus()
'    zlCommFun.OpenIme False
'    Call zl_VsGridLOSTFOCUS(vsPlan)
'End Sub
'
'Private Sub vsPlan_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'    Dim strKey As String, intCol As Integer, strTemp As String
'    Dim str限制项目 As String
'    Dim lng已约数  As Long
'    '数据验证
'    With vsPlan
'        str限制项目 = Switch(Col = 1, "周日", Col = 2, "周一", Col = 3, "周二", Col = 4, "周三", Col = 5, "周四", Col = 6, "周五", True, "周六")
'        strKey = Trim(.EditText): strKey = Replace(strKey, Chr(vbKeyReturn), ""): strKey = Replace(strKey, Chr(10), "")
'        If .Row <= 1 Then Exit Sub
'        If zlCommFun.DblIsValid(strKey, 5, True, False, 0, .ColKey(Col)) = False Then
'            Cancel = True: Exit Sub
'        End If
'        strKey = Format(Abs(Val(strKey)), "####;;;")
'        If Row = 2 Then
'            If mrsRegHistory.RecordCount <> 0 Then
'                mrsRegHistory.Filter = "限制项目='" & str限制项目 & "'"
'                If mrsRegHistory.RecordCount <> 0 Then
'                     lng已约数 = Val(Nvl(mrsRegHistory!统计))
'                     If lng已约数 > Val(strKey) Then
'                        Call MsgBox("限号数小于了已经预约出去的数量[" & lng已约数 & "],不能继续!", vbOKOnly, gstrSysName)
'                        mrsRegHistory.Filter = 0: Cancel = True: Exit Sub
'                      End If
'                End If
'                mrsRegHistory.Filter = 0
'            End If
'            If Val(strKey) < Val(.TextMatrix(3, Col)) Then
'                If MsgBox("限号数小于了限约数,是否清空限约数?", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then Cancel = True: Exit Sub
'                .TextMatrix(3, Col) = ""
'            End If
'        ElseIf Row = 3 Then
'            If mrsRegHistory.RecordCount <> 0 Then
'                mrsRegHistory.Filter = "限制项目='" & str限制项目 & "'"
'                If mrsRegHistory.RecordCount <> 0 Then
'                     lng已约数 = Val(Nvl(mrsRegHistory!统计))
'                     If lng已约数 > Val(strKey) Then
'                        Call MsgBox("限约数小于了已经预约出去的数量[" & lng已约数 & "],不能继续!", vbOKOnly, gstrSysName)
'                        mrsRegHistory.Filter = 0: Cancel = True: Exit Sub
'                      End If
'                End If
'                mrsRegHistory.Filter = 0
'            End If
'
'            If Val(strKey) > Val(.TextMatrix(2, Col)) Then
'                Call MsgBox("限号数小于了限约数,不能继续", vbOKOnly, gstrSysName)
'                Cancel = True: Exit Sub
'            End If
'        End If
'        .EditText = strKey
'    End With
'End Sub
'
'
'
'
'Private Sub cboDoctor_Validate(Cancel As Boolean)
'
'    '指定医生时不能指定多个科室
'    If Trim(cboDoctor.Text) <> "" Then
'        opt分诊(2).Enabled = False
'        opt分诊(3).Enabled = False
'        If opt分诊(2).Value Or opt分诊(3).Value Then opt分诊(0).Value = True
'    Else
'        opt分诊(2).Enabled = True
'        opt分诊(3).Enabled = True
'    End If
'End Sub
'
'Private Sub LoadDoctor()
'    Set mrsDoctor = GetDoctor(Val(cbo科室.ItemData(cbo科室.ListIndex)), "")
'    cboDoctor.Clear
'    Do While Not mrsDoctor.EOF
'        cboDoctor.AddItem mrsDoctor!姓名
'        cboDoctor.ItemData(cboDoctor.NewIndex) = mrsDoctor!ID
'        mrsDoctor.MoveNext
'    Loop
'End Sub
'
'Private Sub cboDoctor_KeyPress(KeyAscii As Integer)
'    Dim lngIdx As Long, lng医生ID As Long
'    If KeyAscii <> 13 Then Exit Sub
'    If cboDoctor.ListIndex <> -1 Then
'        zlCommFun.PressKey vbKeyTab: Exit Sub
'    End If
'    If mrsDoctor Is Nothing Then Exit Sub
'    If Trim(cboDoctor.Text) = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
'
'    If zlPersonSelect(Me, mlngModule, cboDoctor, mrsDoctor, cboDoctor.Text, True, "") = False Then
'        KeyAscii = 0: Exit Sub
'    End If
'    Exit Sub
'End Sub
'
'Private Function Check时段() As Boolean
'    '新增加计划时 获取原有的安排是否具有时段
'    '修改计划时 获取原计划是否具有时段
'   Dim strSQL           As String
'   Dim rsTmp            As ADODB.Recordset
'   If mEditType <> Ed_安排修改 And mEditType <> ed_计划安排 Then Exit Function
'    On Error GoTo Hd
'    If mEditType = ed_计划安排 Then
'        strSQL = " Select 1 As Hdata From 挂号安排时段 Where 安排id =[1] And Rownum=1"
'    Else
'        strSQL = "Select 1  as haveData From 挂号计划时段 Where 计划ID=[2] and Rownum=1"
'    End If
'     Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng安排ID, Val(mstr计划ID))
'     Check时段 = Not rsTmp.EOF
'    Set rsTmp = Nothing
'
'   Exit Function
'Hd:
'   If ErrCenter() = 1 Then
'        Resume
'   End If
'   SaveErrLog
'End Function
'
'
'
'
'
'Private Function LoadRegHistory() As Boolean
'    Dim strSQL As String
'    strSQL = " Select Decode(To_Char(a.发生时间, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五',"
'    strSQL = strSQL & vbCrLf & "                       '7', '周六') As 限制项目, Max(Nvl(a.号序, 0)) As 最大序号, Count(1) As 统计,to_char(Max(发生时间),'hh24:mi:ss') as 发生时间"
'    strSQL = strSQL & vbCrLf & " From 病人挂号记录 a, 挂号安排 b"
'    strSQL = strSQL & vbCrLf & " Where a.记录状态 = 1 And a.发生时间 Between Sysdate And Sysdate + " & IIf(gint预约天数 = 0, 15, gint预约天数) & " And a.号别 = b.号码 And b.Id=[1]"
'    strSQL = strSQL & vbCrLf & " Group By Decode(To_Char(a.发生时间, 'D'), '1', '周日', '2', '周一', '3', '周二', '4', '周三', '5', '周四', '6', '周五',"
'    strSQL = strSQL & vbCrLf & "                             '7', '周六')"
'
'    On Error GoTo Hd:
'    Set mrsRegHistory = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng安排ID)
'    LoadRegHistory = True
'Exit Function
'Hd:
'    If ErrCenter() = 1 Then
'        Resume
'    End If
'    SaveErrLog
'End Function
'
