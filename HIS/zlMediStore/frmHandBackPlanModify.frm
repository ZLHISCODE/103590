VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmHandBackPlanModify 
   Caption         =   "药品退药计划编辑"
   ClientHeight    =   8175
   ClientLeft      =   225
   ClientTop       =   525
   ClientWidth     =   11760
   Icon            =   "frmHandBackPlanModify.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8175
   ScaleWidth      =   11760
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fraControl 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      TabIndex        =   22
      Top             =   7560
      Width           =   13095
      Begin VB.CommandButton cmdClear 
         Caption         =   "清除(&D)"
         Height          =   350
         Left            =   8160
         TabIndex        =   38
         ToolTipText     =   "清除所有行"
         Top             =   120
         Width           =   1100
      End
      Begin VB.TextBox txtFind 
         Height          =   300
         Left            =   2040
         TabIndex        =   27
         Top             =   145
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "保存(&S)"
         Height          =   350
         Left            =   10560
         TabIndex        =   26
         ToolTipText     =   "保存记录"
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton CmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   11760
         TabIndex        =   25
         ToolTipText     =   "不保存退出"
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdReset 
         Caption         =   "重置(&R)"
         Height          =   350
         Left            =   9360
         TabIndex        =   24
         Tag             =   "重新设置退药数量"
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "查找(&F)"
         Height          =   350
         Left            =   105
         TabIndex        =   23
         Top             =   120
         Width           =   1100
      End
      Begin VB.Label lblComment1 
         AutoSize        =   -1  'True
         Caption         =   "在表单内按F3进行连续查找"
         Height          =   180
         Left            =   3840
         TabIndex        =   39
         Top             =   205
         Visible         =   0   'False
         Width           =   2160
      End
      Begin VB.Label lblFindType 
         AutoSize        =   -1  'True
         Caption         =   "编码"
         Height          =   180
         Left            =   1440
         TabIndex        =   28
         Top             =   210
         Visible         =   0   'False
         Width           =   360
      End
   End
   Begin VB.PictureBox picBill 
      Height          =   7455
      Left            =   0
      ScaleHeight     =   7395
      ScaleWidth      =   13035
      TabIndex        =   0
      Top             =   0
      Width           =   13095
      Begin VB.Frame fraComment 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   0
         TabIndex        =   29
         Top             =   6720
         Width           =   12975
         Begin VB.TextBox Txt填制日期 
            Height          =   300
            Left            =   4440
            Locked          =   -1  'True
            TabIndex        =   37
            Top             =   240
            Width           =   1770
         End
         Begin VB.TextBox txt填制人 
            Height          =   300
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   36
            Top             =   240
            Width           =   1050
         End
         Begin VB.TextBox txtNo 
            Height          =   300
            Left            =   360
            Locked          =   -1  'True
            TabIndex        =   35
            Top             =   240
            Width           =   1050
         End
         Begin VB.TextBox txt摘要 
            Height          =   300
            Left            =   7020
            MaxLength       =   40
            TabIndex        =   30
            Top             =   240
            Width           =   5835
         End
         Begin VB.Label Lbl填制人 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "填制人"
            Height          =   180
            Left            =   1680
            TabIndex        =   34
            Top             =   300
            Width           =   540
         End
         Begin VB.Label Lbl填制日期 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "填制日期"
            Height          =   180
            Left            =   3600
            TabIndex        =   33
            Top             =   300
            Width           =   720
         End
         Begin VB.Label lbl摘要 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "摘要"
            Height          =   180
            Left            =   6480
            TabIndex        =   32
            Top             =   300
            Width           =   360
         End
         Begin VB.Label lblNo 
            AutoSize        =   -1  'True
            Caption         =   "NO"
            Height          =   180
            Left            =   120
            TabIndex        =   31
            Top             =   300
            Width           =   180
         End
      End
      Begin VB.Frame fraCondition 
         Height          =   1095
         Left            =   120
         TabIndex        =   1
         Top             =   0
         Width           =   12855
         Begin VB.CommandButton CmdSelecter 
            Caption         =   "…"
            Height          =   300
            Index           =   2
            Left            =   10470
            TabIndex        =   9
            Top             =   580
            Width           =   255
         End
         Begin VB.CommandButton CmdSelecter 
            Caption         =   "…"
            Height          =   300
            Index           =   1
            Left            =   4870
            TabIndex        =   8
            Top             =   600
            Width           =   255
         End
         Begin VB.CommandButton CmdSelecter 
            Caption         =   "…"
            Height          =   300
            Index           =   0
            Left            =   7400
            TabIndex        =   7
            Top             =   180
            Width           =   255
         End
         Begin VB.TextBox txtInput 
            Height          =   300
            Index           =   2
            Left            =   6315
            TabIndex        =   6
            ToolTipText     =   "输入生产商编码、简码或名称"
            Top             =   600
            Width           =   4170
         End
         Begin VB.TextBox txtInput 
            Height          =   300
            Index           =   1
            Left            =   720
            TabIndex        =   5
            ToolTipText     =   "输入供应商编码、简码或名称"
            Top             =   600
            Width           =   4170
         End
         Begin VB.TextBox txtInput 
            Height          =   300
            Index           =   0
            Left            =   4320
            TabIndex        =   4
            ToolTipText     =   "输入药品编码、简码或名称"
            Top             =   180
            Width           =   3090
         End
         Begin VB.CommandButton cmdGet 
            Caption         =   "提取(&G)"
            Height          =   350
            Left            =   11520
            TabIndex        =   3
            Top             =   575
            Width           =   1100
         End
         Begin VB.ComboBox cboStock 
            Height          =   300
            Left            =   720
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   180
            Width           =   2610
         End
         Begin MSComCtl2.DTPicker dtp开始时间 
            Height          =   315
            Left            =   9120
            TabIndex        =   10
            Top             =   180
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   166658051
            CurrentDate     =   36263
         End
         Begin MSComCtl2.DTPicker dtp结束时间 
            Height          =   315
            Left            =   11040
            TabIndex        =   11
            Top             =   180
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "yyyy年MM月dd日"
            Format          =   166658051
            CurrentDate     =   36263
         End
         Begin VB.Label lblInputTxt 
            AutoSize        =   -1  'True
            Caption         =   "生产商"
            Height          =   180
            Index           =   2
            Left            =   5640
            TabIndex        =   20
            Top             =   660
            Width           =   540
         End
         Begin VB.Label lblInputTxt 
            AutoSize        =   -1  'True
            Caption         =   "供应商"
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   19
            Top             =   660
            Width           =   540
         End
         Begin VB.Label lblInputTxt 
            AutoSize        =   -1  'True
            Caption         =   "药品"
            Height          =   180
            Index           =   0
            Left            =   3840
            TabIndex        =   18
            Top             =   240
            Width           =   360
         End
         Begin VB.Label lbl至 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "～"
            Height          =   180
            Left            =   10800
            TabIndex        =   17
            Top             =   240
            Width           =   180
         End
         Begin VB.Label lbl时间 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "入库日期"
            Height          =   180
            Left            =   8280
            TabIndex        =   16
            Top             =   240
            Width           =   720
         End
         Begin VB.Label LblStock 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "库房"
            Height          =   180
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   360
         End
         Begin VB.Label lblFlag 
            AutoSize        =   -1  'True
            Caption         =   "*"
            Height          =   180
            Index           =   0
            Left            =   3360
            TabIndex        =   14
            Top             =   240
            Width           =   90
         End
         Begin VB.Label lblFlag 
            AutoSize        =   -1  'True
            Caption         =   "*"
            Height          =   180
            Index           =   1
            Left            =   7700
            TabIndex        =   13
            Top             =   240
            Width           =   90
         End
         Begin VB.Label lblFlag 
            AutoSize        =   -1  'True
            Caption         =   "*"
            Height          =   180
            Index           =   2
            Left            =   12660
            TabIndex        =   12
            Top             =   240
            Width           =   90
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfBill 
         Height          =   3255
         Left            =   120
         TabIndex        =   21
         Top             =   1200
         Width           =   4335
         _cx             =   7646
         _cy             =   5741
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
         BackColorAlternate=   15724527
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmHandBackPlanModify.frx":038A
         ScrollTrack     =   0   'False
         ScrollBars      =   3
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
End
Attribute VB_Name = "frmHandBackPlanModify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlng库房ID As Long
Private mintUnit As Integer
Private mstrNo As String
Private mblnSuccess As Boolean
Private Const MStrCaption As String = "药品退药计划编辑"

Dim mlngFind As Long                            '用于查找
Dim mrsFindName As ADODB.Recordset              '用于查找

Private Enum InputType
    药品 = 0
    供应商 = 1
    生产商 = 2
End Enum

'汇总，明细列表标题
Private Const mconstBillHead = "药品ID,1,0|供应商ID,1,0|序号,4,500|供应商,1,2500|药品编码,1,1000|药品名称,1,2000|商品名,1,2000|规格,1,2000|生产商,1,2000|批号,1,1000|效期,1,1000|单位,1,800|退药数量,7,1000|成本价,7,1000|成本金额,7,1000|包装,7,0"

Private Enum 明细列表
    药品id = 0
    供应商id = 1
    序号 = 2
    供应商 = 3
    药品编码 = 4
    药品名称 = 5
    商品名 = 6
    规格 = 7
    生产商 = 8
    批号 = 9
    效期 = 10
    单位 = 11
    数量 = 12
    成本价 = 13
    成本金额 = 14
    包装 = 15
    列数 = 16
End Enum

Private Function CheckRepeat(ByVal strInfo As String, Optional ByVal intExceptCol As Integer = 0) As Boolean
    '检查是否重复
    '检查规则：药品ID、供应商ID、生产商、批号
    'strInfo格式：药品ID;供应商ID;生产商;批号
    'intExceptCol：排除该列
    'CheckRepeat返回：True-重复;False-不重复
    
    Dim lng药品ID As Long
    Dim lng供应商ID As Long
    Dim str生产商 As String
    Dim str批号 As String
    Dim i As Integer
    
    If vsfBill.rows = 1 Then Exit Function
    If vsfBill.TextMatrix(1, 明细列表.药品id) = "" Then Exit Function
    
    lng药品ID = Split(strInfo, ";")(0)
    lng供应商ID = Split(strInfo, ";")(1)
    str生产商 = Split(strInfo, ";")(2)
    str批号 = Split(strInfo, ";")(3)
    
    With vsfBill
        For i = 1 To .rows - 1
            If i <> intExceptCol And Val(.TextMatrix(i, 明细列表.药品id)) = lng药品ID And Val(.TextMatrix(i, 明细列表.供应商id)) = lng供应商ID _
                And .TextMatrix(i, 明细列表.生产商) = str生产商 And .TextMatrix(i, 明细列表.批号) = str批号 Then
                CheckRepeat = True
                Exit Function
            End If
        Next
    End With
End Function
Private Sub IniGrid()
    Dim i As Integer
    Dim strArr As Variant
    Dim strTemp As Variant
    
    strTemp = Split(mconstBillHead, "|")
    With vsfBill
        .Redraw = flexRDNone
        .rows = 1
        .Cols = 明细列表.列数
        .SelectionMode = flexSelectionByRow
        .ExplorerBar = flexExSortShow
        For i = 0 To .Cols - 1
            strArr = Split(strTemp(i), ",")
            .TextMatrix(0, i) = strArr(0)
            .ColAlignment(i) = strArr(1)
            .ColWidth(i) = strArr(2)
            .FixedAlignment(i) = flexAlignCenterCenter
        Next

        .Redraw = flexRDDirect
    End With
End Sub
Private Sub GetDate(ByVal strNo As String)
    '提取已存在的单据明细
    Dim rsTmp As ADODB.Recordset
    Dim strSubUnit As String
    
    If strNo = "" Then Exit Sub
    On Error GoTo errHandle
    '单位，包装换算
    '单位系数：1-售价;2-门诊;3-住院;4-药库
    Select Case mintUnit
    Case 1
        strSubUnit = "D.计算单位 单位,1 包装 "
    Case 2
        strSubUnit = "B.门诊单位 单位,B.门诊包装 包装 "
    Case 3
        strSubUnit = "B.住院单位 单位,B.住院包装 包装 "
    Case 4
        strSubUnit = "B.药库单位 单位,B.药库包装 包装 "
    End Select
    
    gstrSQL = "Select Distinct A.序号, A.药品id, D.编码 As 药品编码,D.名称 As 通用名,E.名称 As 商品名, " & _
        " D.规格, A.实际数量, A.效期,A.成本价, A.成本金额, A.产地 As 生产商, A.批号,A.供药单位id,F.名称 As 供应商, " & _
        " A.填制人, A.填制日期, A.摘要, " & strSubUnit & _
        " From 药品退药计划 A, 药品规格 B, 收费项目目录 D, 收费项目别名 E, 供应商 F " & _
        " Where A.药品id = B.药品id And B.药品id = D.ID And B.药品id = E.收费细目id(+) And E.性质(+) = 3 And A.供药单位id = F.ID And A.No = [1] " & _
        " Order By A.序号 "
    Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, "提取入库信息", strNo)
    
    vsfBill.rows = 1
    
    If rsTmp.EOF Then Exit Sub
    
    With rsTmp
        txtNo.Text = strNo
        txt填制人.Text = !填制人
        Txt填制日期.Text = Format(!填制日期, "yyyy-mm-dd hh:mm:ss")
        txt摘要.Text = Nvl(!摘要)
        Do While Not .EOF
            vsfBill.rows = vsfBill.rows + 1
            
            vsfBill.TextMatrix(vsfBill.rows - 1, 明细列表.序号) = .AbsolutePosition
            
            vsfBill.TextMatrix(vsfBill.rows - 1, 明细列表.药品id) = !药品id
            vsfBill.TextMatrix(vsfBill.rows - 1, 明细列表.供应商id) = !供药单位ID
            vsfBill.TextMatrix(vsfBill.rows - 1, 明细列表.药品编码) = !药品编码
            If gint药品名称显示 = 0 Or gint药品名称显示 = 2 Then
                vsfBill.TextMatrix(vsfBill.rows - 1, 明细列表.药品名称) = !通用名
            Else
                vsfBill.TextMatrix(vsfBill.rows - 1, 明细列表.药品名称) = IIf(IsNull(!商品名), !通用名, !商品名)
            End If
            vsfBill.TextMatrix(vsfBill.rows - 1, 明细列表.商品名) = IIf(IsNull(!商品名), "", !商品名)
            vsfBill.TextMatrix(vsfBill.rows - 1, 明细列表.供应商) = Nvl(!供应商)
            vsfBill.TextMatrix(vsfBill.rows - 1, 明细列表.规格) = Nvl(!规格)
            vsfBill.TextMatrix(vsfBill.rows - 1, 明细列表.单位) = Nvl(!单位)
            vsfBill.TextMatrix(vsfBill.rows - 1, 明细列表.数量) = zlStr.FormatEx(!实际数量 / !包装, 2, , True)
            vsfBill.TextMatrix(vsfBill.rows - 1, 明细列表.成本价) = zlStr.FormatEx(!成本价 * !包装, 5, , True)
            vsfBill.TextMatrix(vsfBill.rows - 1, 明细列表.成本金额) = zlStr.FormatEx(!成本金额, 2, , True)
            vsfBill.TextMatrix(vsfBill.rows - 1, 明细列表.生产商) = Nvl(!生产商)
            vsfBill.TextMatrix(vsfBill.rows - 1, 明细列表.批号) = Nvl(!批号)
            vsfBill.TextMatrix(vsfBill.rows - 1, 明细列表.包装) = !包装
            
            vsfBill.TextMatrix(vsfBill.rows - 1, 明细列表.效期) = Format(IIf(IsNull(!效期), "", !效期), "yyyy-mm-dd")
                    
            If gtype_UserSysParms.P149_效期显示方式 = 1 And vsfBill.TextMatrix(vsfBill.rows - 1, 明细列表.效期) <> "" Then
                '换算为有效期
                vsfBill.TextMatrix(vsfBill.rows - 1, 明细列表.效期) = Format(DateAdd("D", -1, vsfBill.TextMatrix(vsfBill.rows - 1, 明细列表.效期)), "yyyy-mm-dd")
            End If
            
            .MoveNext
        Loop
        
        vsfBill.Cell(flexcpForeColor, 1, 明细列表.数量, vsfBill.rows - 1, 明细列表.数量) = vbBlue
        vsfBill.Cell(flexcpFontBold, 1, 明细列表.数量, vsfBill.rows - 1, 明细列表.数量) = True
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetNewDate()
    '产生新数据
    Dim lng库房ID As Long
    Dim lng药品ID As Long
    Dim str开始时间 As String
    Dim str结束时间 As String
    Dim lng供应商ID As Long
    Dim str生产商 As String
    
    Dim rsTmp As ADODB.Recordset
    Dim strSubUnit As String
    Dim strSqlCondition As String
    
    On Error GoTo errHandle
    lng库房ID = Val(cboStock.ItemData(cboStock.ListIndex))
    lng药品ID = Val(txtInput(InputType.药品).Tag)
    str开始时间 = Format(dtp开始时间.Value, "YYYY-MM-DD") & " 00:00:01"
    str结束时间 = Format(dtp结束时间.Value, "YYYY-MM-DD") & " 23:59:59"
    If Val(txtInput(InputType.供应商).Tag) > 0 And Trim(txtInput(InputType.供应商).Text) <> "" Then
        lng供应商ID = Val(txtInput(InputType.供应商).Tag)
    End If
    str生产商 = Trim(txtInput(InputType.生产商).Text)
    
    If lng库房ID = 0 Or lng药品ID = 0 Then Exit Sub
        
    strSqlCondition = " And A.库房id + 0 = [1] And A.药品id + 0 = [2] And A.审核日期 Between [3] And [4] "
    
    If lng供应商ID > 0 Then
        strSqlCondition = strSqlCondition & " And A.供药单位id = [5] "
    End If
    
    If str生产商 <> "" Then
        strSqlCondition = strSqlCondition & " And A.产地 = [6] "
    End If
        
    '单位，包装换算
    '单位系数：1-售价;2-门诊;3-住院;4-药库
    Select Case mintUnit
    Case 1
        strSubUnit = "D.计算单位 单位,1 包装 "
    Case 2
        strSubUnit = "B.门诊单位 单位,B.门诊包装 包装 "
    Case 3
        strSubUnit = "B.住院单位 单位,B.住院包装 包装 "
    Case 4
        strSubUnit = "B.药库单位 单位,B.药库包装 包装 "
    End Select
    
    gstrSQL = "Select Distinct A.药品id, D.编码 As 药品编码,D.名称 As 通用名,E.名称 As 商品名, " & _
        " D.规格, A.效期, A.实际数量, A.成本价, A.成本金额, A.产地 As 生产商, A.批号,A.供药单位id,F.名称 As 供应商, " & strSubUnit & _
        " From 药品收发记录 A, 药品规格 B, 收费项目目录 D, 收费项目别名 E, 供应商 F " & _
        " Where A.药品id = B.药品id And B.药品id = D.ID And B.药品id = E.收费细目id(+) And E.性质(+) = 3 And A.供药单位id = F.ID And A.单据 = 1 " & strSqlCondition & _
        " Order By F.名称"
    Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, "提取入库信息", lng库房ID, lng药品ID, CDate(str开始时间), CDate(str结束时间), lng供应商ID, str生产商)
    
    If rsTmp.EOF Then Exit Sub
    With rsTmp
        Do While Not .EOF
            '检查是否重复
            If CheckRepeat(!药品id & ";" & !供药单位ID & ";" & !生产商 & ";" & !批号) = False Then
                vsfBill.rows = vsfBill.rows + 1
                
                vsfBill.TextMatrix(vsfBill.rows - 1, 明细列表.药品id) = !药品id
                vsfBill.TextMatrix(vsfBill.rows - 1, 明细列表.供应商id) = !供药单位ID
                vsfBill.TextMatrix(vsfBill.rows - 1, 明细列表.序号) = vsfBill.rows - 1
                vsfBill.TextMatrix(vsfBill.rows - 1, 明细列表.供应商) = !供应商
                vsfBill.TextMatrix(vsfBill.rows - 1, 明细列表.药品编码) = !药品编码
                If gint药品名称显示 = 0 Or gint药品名称显示 = 2 Then
                    vsfBill.TextMatrix(vsfBill.rows - 1, 明细列表.药品名称) = !通用名
                Else
                    vsfBill.TextMatrix(vsfBill.rows - 1, 明细列表.药品名称) = IIf(IsNull(!商品名), !通用名, !商品名)
                End If
                vsfBill.TextMatrix(vsfBill.rows - 1, 明细列表.商品名) = IIf(IsNull(!商品名), "", !商品名)
                vsfBill.TextMatrix(vsfBill.rows - 1, 明细列表.规格) = Nvl(!规格)
                vsfBill.TextMatrix(vsfBill.rows - 1, 明细列表.生产商) = Nvl(!生产商)
                vsfBill.TextMatrix(vsfBill.rows - 1, 明细列表.批号) = Nvl(!批号)
                vsfBill.TextMatrix(vsfBill.rows - 1, 明细列表.效期) = Format(IIf(IsNull(!效期), "", !效期), "yyyy-mm-dd")
                    
                If gtype_UserSysParms.P149_效期显示方式 = 1 And vsfBill.TextMatrix(vsfBill.rows - 1, 明细列表.效期) <> "" Then
                    '换算为有效期
                    vsfBill.TextMatrix(vsfBill.rows - 1, 明细列表.效期) = Format(DateAdd("D", -1, vsfBill.TextMatrix(vsfBill.rows - 1, 明细列表.效期)), "yyyy-mm-dd")
                End If
                
                vsfBill.TextMatrix(vsfBill.rows - 1, 明细列表.单位) = Nvl(!单位)
                vsfBill.TextMatrix(vsfBill.rows - 1, 明细列表.数量) = ""
                vsfBill.TextMatrix(vsfBill.rows - 1, 明细列表.成本价) = zlStr.FormatEx(!成本价 * !包装, 5, , True)
                vsfBill.TextMatrix(vsfBill.rows - 1, 明细列表.成本金额) = ""
                
                
                vsfBill.TextMatrix(vsfBill.rows - 1, 明细列表.包装) = !包装
            End If
            
            .MoveNext
        Loop
        vsfBill.Cell(flexcpForeColor, 1, 明细列表.数量, vsfBill.rows - 1, 明细列表.数量) = vbBlue
        vsfBill.Cell(flexcpFontBold, 1, 明细列表.数量, vsfBill.rows - 1, 明细列表.数量) = True
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadStock()
    '取库房：只取药库属性的库房
    
    Dim rsTmp As ADODB.Recordset
    Dim lngDrugStoreIndex As Long
    Dim intIndex As Integer
    
    On Error GoTo errHandle
    gstrSQL = "Select Distinct A.ID, A.名称 " & _
              "From 部门性质说明 C, 部门性质分类 B, 部门表 A " & _
              "Where (A.站点 = [1] Or A.站点 is Null) And C.工作性质 = B.名称 And Instr('HIJ', B.编码, 1) > 0 " & _
              "  And A.ID = C.部门id And To_Char(A.撤档时间, 'yyyy-MM-dd') = '3000-01-01' " & _
              "Order By A.名称 "
    Set rsTmp = zlDataBase.OpenSQLRecord(gstrSQL, "取所有药库属性的库房", gstrNodeNo)
    
    If rsTmp.EOF Then
        MsgBox "至少应该设置一个具有药库性质的部门,请查看部门管理！", vbInformation, gstrSysName
        Unload Me
        Exit Sub
    End If
    
    With rsTmp
        cboStock.Clear
        
        Do While Not .EOF
            cboStock.AddItem !名称
            cboStock.ItemData(cboStock.NewIndex) = !id
            If !id = mlng库房ID Then
                lngDrugStoreIndex = intIndex
            End If
            intIndex = intIndex + 1
            .MoveNext
        Loop
        
        cboStock.ListIndex = lngDrugStoreIndex
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub RefreshSerialNumber()
    '重新调整序号：表单排序、删除行后使用
    
    Dim i As Integer
        
    With vsfBill
        If .rows = 2 Then Exit Sub
        For i = 1 To .rows - 1
            .TextMatrix(i, 明细列表.序号) = i
        Next
    End With
End Sub

Private Function SelectInput(ByVal intType As Integer, ByVal strkey As String, ByVal sngX As Single, ByVal sngY As Single, ByVal sngH As Single) As String
    '选择器：支持对药品、供应商、生产商的选择
    'intType：0-药品;1-供应商;2-生产商
    'strKey：空-全部;非空-模糊匹配
    'SelectInput返回值：空-没找到匹配记录;
    '                 非空-药品（药品ID;药品名称;规格;单位;包装）
    '                     -供应商（供应商ID;供应商名称）
    '                     -生产商（生产商ID;生产商名称）
    
    Dim rsTemp As ADODB.Recordset
    Dim blnCancel As Boolean
    Dim strSubUnit As String
    Dim strFindString As String
    Dim strReturn As String
    Dim strSql药品 As String
    
    Err = 0: On Error GoTo ErrHand:
    
    strkey = UCase(Trim(strkey))
    
    Select Case intType
    Case InputType.药品
        '单位，包装换算
        '单位系数：1-售价;2-门诊;3-住院;4-药库
        Select Case mintUnit
        Case 1
            strSubUnit = "B.计算单位 单位,1 包装 "
        Case 2
            strSubUnit = "A.门诊单位 单位,A.门诊包装 包装 "
        Case 3
            strSubUnit = "A.住院单位 单位,A.住院包装 包装 "
        Case 4
            strSubUnit = "A.药库单位 单位,A.药库包装 包装 "
        End Select
        
        If strkey <> "" Then
            strFindString = " And (B.编码 Like [1] OR B.名称 Like [2] OR C.简码 LIKE [2])"
            
            If IsNumeric(strkey) Then                         '10,11.输入全是数字时只匹配编码'对于药品,则要匹配简码(码类为3的数字码)
                If Mid(gtype_UserSysParms.P44_输入匹配, 1, 1) = "1" Then strFindString = " And (B.编码 Like [1] Or B.简码 Like [2] And C.码类=3)"
            ElseIf zlStr.IsCharAlpha(strkey) Then         '01,11.输入全是字母时只匹配简码
                If Mid(gtype_UserSysParms.P44_输入匹配, 2, 1) = "1" Then strFindString = " And C.简码 Like [2] "
            ElseIf zlStr.IsCharChinese(strkey) Then
                strFindString = " And B.名称 Like [2] "
            End If
        End If
        
        If strkey = "" Then
            If gint药品名称显示 = 0 Then
                strSql药品 = ",'['||编码||']'|| 通用名 As 药品名称"
            ElseIf gint药品名称显示 = 1 Then
                strSql药品 = ",'['||编码||']'|| Nvl(商品名,通用名) As 药品名称"
            ElseIf gint药品名称显示 = 2 Then
                strSql药品 = ",'['||编码||']'|| 通用名 As 药品名称,商品名"
            End If
            
            gstrSQL = "Select Rownum As ID, 药品id " & strSql药品 & ", 规格, 产地 as 生产商, 单位, 包装,商品名 " & _
                " From (Select Distinct A.药品id, B.编码, B.名称 As 通用名, C.名称 As 商品名, B.规格,B.产地,  " & strSubUnit & _
                " From 药品规格 A, " & _
                " (Select B.ID, B.编码, B.名称, B.规格,B.产地,B.计算单位 From 收费项目目录 B, 收费项目别名 C " & _
                " Where (B.站点 = [3] Or B.站点 is Null) And B.ID = C.收费细目id And B.类别 In ('5', '6', '7') " & strFindString & ") B, 收费项目别名 C " & _
                " Where A.药品id = B.ID And A.药品id = C.收费细目id(+) And C.性质(+) = 3 "
            gstrSQL = gstrSQL & " Order By B.编码)"
        Else
            strSql药品 = ",'['||编码||']'|| 输入名称 As 药品名称"
            
            gstrSQL = "Select Rownum As ID, 药品id " & strSql药品 & ", 规格, 产地 as 生产商, 单位, 包装,商品名 " & _
                " From (Select Distinct A.药品id, B.编码, B.输入名称, B.名称 As 通用名, C.名称 As 商品名, B.规格,B.产地,  " & strSubUnit & _
                " From 药品规格 A, " & _
                " (Select B.ID, B.编码, B.名称, B.规格,B.产地,B.计算单位, C.名称 As 输入名称 From 收费项目目录 B, 收费项目别名 C " & _
                " Where (B.站点 = [3] Or B.站点 is Null) And B.ID = C.收费细目id And B.类别 In ('5', '6', '7') " & strFindString & ") B, 收费项目别名 C " & _
                " Where A.药品id = B.ID And A.药品id = C.收费细目id(+) And C.性质(+) = 3 "
            gstrSQL = gstrSQL & " Order By B.编码)"
        End If
        
        Set rsTemp = zlDataBase.ShowSQLSelect(Me, gstrSQL, 0, "药品选择器", False, "", "选择药品", False, False, True, sngX, sngY, sngH, blnCancel, False, False, _
                        strkey & "%", "%" & strkey & "%", _
                        gstrNodeNo)
        
        If blnCancel = True Then Exit Function
        
        If rsTemp Is Nothing Then
            strReturn = ""
        Else
            strReturn = rsTemp!药品id & ";" & rsTemp!药品名称 & ";" & rsTemp!商品名 & ";" & rsTemp!规格 & ";" & rsTemp!单位 & ";" & rsTemp!包装
        End If
    Case InputType.供应商
        gstrSQL = "Select id,名称,编码,简码 From 供应商 " & _
                  "Where (站点 = [3] Or 站点 is Null) " & _
                  "  And (To_Char(撤档时间,'yyyy-MM-dd')='3000-01-01' or 撤档时间 is null ) And 末级=1 And (substr(类型,1,1)=1 Or Nvl(末级,0)=0) " & _
                  "  And (编码 like [1] or 简码 like [2] or 名称 like [2])"
        Set rsTemp = zlDataBase.ShowSQLSelect(Me, gstrSQL, 0, "供应商选择器", False, "", "选择供应商", False, False, True, sngX, sngY, sngH, blnCancel, False, False, _
                        strkey & "%", "%" & strkey & "%", _
                        gstrNodeNo)
        
        If blnCancel = True Then Exit Function
        
        If rsTemp Is Nothing Then
            MsgBox "输入值无效！", vbInformation, gstrSysName
            strReturn = ""
        Else
            strReturn = rsTemp!id & ";" & rsTemp!名称
        End If
    Case InputType.生产商
        gstrSQL = "Select Rownum As ID,名称,编码,简码 From 药品生产商 " & _
                  "Where (站点 = [3] Or 站点 is Null) And (编码 like [1] Or 简码 like [2] Or 名称 like [2]) Order By 编码"
        Set rsTemp = zlDataBase.ShowSQLSelect(Me, gstrSQL, 0, "生产商选择器", False, "", "选择生产商", False, False, True, sngX, sngY, sngH, blnCancel, False, False, _
                        strkey & "%", "%" & strkey & "%", _
                        gstrNodeNo)
        
        If blnCancel = True Then Exit Function
        
        If rsTemp Is Nothing Then
            MsgBox "输入值无效！", vbInformation, gstrSysName
            strReturn = ""
        Else
            strReturn = rsTemp!id & ";" & rsTemp!名称
        End If
    End Select
    
    SelectInput = strReturn
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub FindGridRow(ByVal strInput As String)
    Dim lngStart As Long
    Dim str编码 As String, str名称 As String, str简码 As String
    Dim str其他名称 As String
    Dim blnEnd As Boolean
    Dim lngFindRow As Long
    Dim str药名 As String
    
    '查找药品
    On Error GoTo errHandle
    If strInput <> txtFind.Tag Then
        '表示新的查找
        If vsfBill.rows > 1 Then vsfBill.Row = 1
        txtFind.Tag = strInput
        
        gstrSQL = "Select Distinct A.Id,'[' || A.编码 || ']' As 药品编码, A.名称 As 通用名, B.名称 As 商品名 " & _
                  "From 收费项目目录 A,收费项目别名 B " & _
                  "Where (A.站点 = [3] Or A.站点 is Null) And A.Id =B.收费细目id And A.类别 In ('5','6','7') " & _
                  "  And (A.编码 Like [1] Or B.名称 Like [2] Or B.简码 Like [2] )" & _
                  "Order By '[' || A.编码 || ']' "
        Set mrsFindName = zlDataBase.OpenSQLRecord(gstrSQL, "取匹配的药品ID", strInput & "%", "%" & strInput & "%", gstrNodeNo)
        
        If mrsFindName.RecordCount = 0 Then Exit Sub
    End If
    
    
    '开始查找
    If mrsFindName.State <> adStateOpen Then Exit Sub
    If mrsFindName.RecordCount = 0 Then Exit Sub

    If mrsFindName.EOF Then mrsFindName.MoveFirst
    Do While Not mrsFindName.EOF
        str药名 = mrsFindName!药品编码
        str药名 = Mid(str药名, 2, Len(str药名) - 2)
        lngFindRow = vsfBill.FindRow(str药名, lngStart, 明细列表.药品编码, True, True)
        If lngFindRow > 0 Then
            vsfBill.Select lngFindRow, 1, lngFindRow, vsfBill.Cols - 1
            vsfBill.TopRow = lngFindRow
            mlngFind = lngFindRow
            mrsFindName.MoveNext
            If lngStart >= vsfBill.rows - 1 Then
                lngStart = 1
            Else
                lngStart = lngStart + 1
            End If
            Exit Do
        End If
        mrsFindName.MoveNext
        If mrsFindName.EOF Then
            MsgBox "查找已至底，继续查找将从顶开始！", vbInformation, gstrSysName
        End If
    Loop
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Public Sub ShowForm(FrmMain As Form, ByVal lng库房ID As Long, ByVal intUnit As Integer, ByRef BlnSuccess As Boolean, Optional ByVal strNo As String = "")
    mlng库房ID = lng库房ID
    mintUnit = intUnit
    mstrNo = strNo
    mblnSuccess = False
    
    Me.Show vbModal, FrmMain
    
    BlnSuccess = mblnSuccess
End Sub

Private Function GetSeleterReturn(ByVal intType As Integer, ByVal objInputObj As Object, ByVal strInput As String) As Boolean
    '通过输入项目返回选择器的选择值
    'intType：0-药品;1-供应商;2-生产商
    'objInputObj：两种类型的录入对象，TextBox和VSFlexGrid
    'strInput：输入值，可以是编码、简码、名称
    Dim vRect As RECT
    Dim strReturn As String
    Dim sngX As Single
    Dim sngY As Single
    Dim sngH As Single
    
    '根据对象的类型取对象的参照位置
    If TypeName(objInputObj) = "TextBox" Then
        vRect = zlControl.GetControlRect(objInputObj.hWnd)
        sngX = vRect.Left
        sngY = vRect.Top
        sngH = objInputObj.Height
    ElseIf TypeName(objInputObj) = "VSFlexGrid" Then
        Call CalcPosition(sngX, sngY, objInputObj)
        sngY = sngY - objInputObj.CellHeight
        sngH = objInputObj.CellHeight
    Else
        Exit Function
    End If
    
    '得到选择器的返回值
    strReturn = SelectInput(intType, strInput, sngX, sngY, sngH)
    
    '根据实际业务处理
    If TypeName(objInputObj) = "TextBox" Then
'        strReturn="ID;名称"
        If strReturn = "" Then Exit Function
            
        objInputObj.Tag = Val(Split(strReturn, ";")(0))
        objInputObj.Text = Split(strReturn, ";")(1)
    Else
        If strReturn = "" Then
            Select Case intType
            Case InputType.药品
                If Val(objInputObj.TextMatrix(objInputObj.Row, 明细列表.药品id)) > 0 Then
                    Exit Function
                End If
            Case InputType.供应商
                If Val(objInputObj.TextMatrix(objInputObj.Row, 明细列表.供应商id)) > 0 Then
                    Exit Function
                End If
            Case InputType.生产商
                If Trim(objInputObj.TextMatrix(objInputObj.Row, 明细列表.生产商)) <> "" Then
                    Exit Function
                End If
            End Select
            
            objInputObj.TextMatrix(objInputObj.Row, objInputObj.Col) = objInputObj.EditText
            objInputObj.Cell(flexcpText, objInputObj.Row, objInputObj.Col) = ""
            Exit Function
        Else
            Select Case intType
            Case InputType.药品
        '        strReturn="药品ID;药品名称;规格;单位;包装"
                objInputObj.TextMatrix(objInputObj.Row, 明细列表.药品id) = Val(Split(strReturn, ";")(0))
                objInputObj.TextMatrix(objInputObj.Row, 明细列表.药品名称) = Split(strReturn, ";")(1)
                objInputObj.TextMatrix(objInputObj.Row, 明细列表.商品名) = Split(strReturn, ";")(2)
                objInputObj.TextMatrix(objInputObj.Row, 明细列表.规格) = Split(strReturn, ";")(3)
                objInputObj.TextMatrix(objInputObj.Row, 明细列表.单位) = Split(strReturn, ";")(4)
                objInputObj.TextMatrix(objInputObj.Row, 明细列表.包装) = Val(Split(strReturn, ";")(5))
                
                objInputObj.EditText = Split(strReturn, ";")(1)
            Case InputType.供应商
        '        strReturn="供应商ID;供应商名称"
                objInputObj.TextMatrix(objInputObj.Row, 明细列表.供应商id) = Val(Split(strReturn, ";")(0))
                objInputObj.TextMatrix(objInputObj.Row, 明细列表.供应商) = Split(strReturn, ";")(1)
                
                objInputObj.EditText = Split(strReturn, ";")(1)
            Case InputType.生产商
        '        strReturn="生产商ID;生产商名称"
                objInputObj.TextMatrix(objInputObj.Row, 明细列表.生产商) = Split(strReturn, ";")(1)
                
                objInputObj.EditText = Split(strReturn, ";")(1)
            End Select
        End If
    End If
    
    GetSeleterReturn = True
End Function

Private Sub cboStock_Click()
    Call SetSelectorRS(1, "药品外购入库管理", cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex))
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdClear_Click()
    '清除所有行
    With vsfBill
        .rows = 1
    End With
End Sub

Private Sub cmdFind_Click()
    Dim blnVisible As Boolean
    '查找或查找下一条
    blnVisible = lblFindType.Visible Xor True
    lblFindType.Visible = blnVisible
    txtFind.Visible = blnVisible
    lblComment1.Visible = blnVisible
    If blnVisible Then txtFind.SetFocus
End Sub

Private Sub cmdGet_Click()
    Call GetNewDate
End Sub
Private Sub cmdReset_Click()
    '清空退药数量和成本金额
    With vsfBill
        If .rows = 1 Then Exit Sub
        If .TextMatrix(1, 0) = "" Then Exit Sub
        .Cell(flexcpText, 1, 明细列表.数量, .rows - 1, 明细列表.数量) = ""
        .Cell(flexcpText, 1, 明细列表.成本金额, .rows - 1, 明细列表.成本金额) = ""
    End With
End Sub

Private Sub CmdSave_Click()
    Dim strNo_In As String
    Dim int序号_In As Integer
    Dim lng药品id_In As Long
    Dim lng供药单位id_In As Long
    Dim dbl数量_In As Double
    Dim dbl成本价_In As Double
    Dim dbl成本金额_In As Double
    Dim str产地_In As String
    Dim str批号_In As String
    Dim str填制人_In As String
    Dim str填制日期_In As String
    Dim str摘要_In As String
    Dim str效期 As String
    
    Dim blnTrans As Boolean
    Dim i As Integer
    Dim intCount As Integer
    Dim arrSql As Variant
    
    On Error GoTo errHandle
    
    arrSql = Array()
    '如果是已有单据，先删除
    strNo_In = Trim(txtNo.Text)
    If strNo_In <> "" Then
'        gcnOracle.BeginTrans
        blnTrans = True
    
        gstrSQL = "Zl_药品退药计划_Delete('" & strNo_In & "')"
        ReDim Preserve arrSql(UBound(arrSql) + 1)
        arrSql(UBound(arrSql)) = gstrSQL
    End If
    
    '保存新单据
    With vsfBill
        '产生新NO号
        If strNo_In = "" Then
            strNo_In = Sys.GetNextNo(100)
        End If
        str填制人_In = txt填制人.Text
        str填制日期_In = Format(Sys.Currentdate, "yyyy-mm-dd hh:mm:ss")
        str摘要_In = txt摘要.Text
        
        For i = 1 To .rows - 1
            If Val(.TextMatrix(i, 明细列表.数量)) > 0 Then
                int序号_In = intCount + 1
                lng药品id_In = Val(.TextMatrix(i, 明细列表.药品id))
                lng供药单位id_In = Val(.TextMatrix(i, 明细列表.供应商id))
                dbl数量_In = zlStr.FormatEx(Val(.TextMatrix(i, 明细列表.数量)) * Val(.TextMatrix(i, 明细列表.包装)), 5, , True)
                dbl成本价_In = zlStr.FormatEx(Val(.TextMatrix(i, 明细列表.成本价)) / Val(.TextMatrix(i, 明细列表.包装)), 5, , True)
                dbl成本金额_In = Val(.TextMatrix(i, 明细列表.成本金额))
                str产地_In = IIf(Trim(.TextMatrix(i, 明细列表.生产商)) = "", "", .TextMatrix(i, 明细列表.生产商))
                str批号_In = IIf(Trim(.TextMatrix(i, 明细列表.批号)) = "", "", .TextMatrix(i, 明细列表.批号))
                
                str效期 = IIf(Trim(.TextMatrix(i, 明细列表.效期)) = "", "", .TextMatrix(i, 明细列表.效期))
                
                If gtype_UserSysParms.P149_效期显示方式 = 1 And str效期 <> "" Then
                    '换算为失效期来保存
                    str效期 = Format(DateAdd("D", 1, str效期), "yyyy-mm-dd")
                End If
                
                gstrSQL = "Zl_药品退药计划_Insert("
                'NO
                gstrSQL = gstrSQL & "'" & strNo_In & "'"
                '序号
                gstrSQL = gstrSQL & "," & int序号_In
                '药品ID
                gstrSQL = gstrSQL & "," & lng药品id_In
                '供药单位ID
                gstrSQL = gstrSQL & "," & lng供药单位id_In
                '退药数量
                gstrSQL = gstrSQL & "," & dbl数量_In
                '成本价
                gstrSQL = gstrSQL & "," & dbl成本价_In
                '成本金额
                gstrSQL = gstrSQL & "," & dbl成本金额_In
                '产地
                gstrSQL = gstrSQL & ",'" & str产地_In & "'"
                '批号
                gstrSQL = gstrSQL & ",'" & str批号_In & "'"
                '效期
                gstrSQL = gstrSQL & "," & IIf(str效期 = "", "Null", "to_date('" & Format(str效期, "yyyy-mm-dd") & "','yyyy-mm-dd')")
                '填制人
                gstrSQL = gstrSQL & ",'" & str填制人_In & "'"
                '填制日期
                gstrSQL = gstrSQL & ",to_date('" & str填制日期_In & "','yyyy-mm-dd HH24:MI:SS')"
                '摘要
                gstrSQL = gstrSQL & ",'" & str摘要_In & "'"
                gstrSQL = gstrSQL & ")"

                ReDim Preserve arrSql(UBound(arrSql) + 1)
                arrSql(UBound(arrSql)) = gstrSQL
                
                intCount = intCount + 1
            End If
        Next
    End With
    
    If intCount = 0 Then
        MsgBox "请录入退药数量！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    gcnOracle.BeginTrans
    For i = 0 To UBound(arrSql)
        Call zlDataBase.ExecuteProcedure(CStr(arrSql(i)), "SaveCard")
    Next
    gcnOracle.CommitTrans
    mblnSuccess = True
    
    Unload Me
    Exit Sub
errHandle:
    If blnTrans = True Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub CmdSelecter_Click(Index As Integer)
    Dim RecReturn As ADODB.Recordset
    
    If grsMaster.State = adStateClosed Then
        Call SetSelectorRS(1, "药品外购入库管理", cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex))
    End If
    
    If Index = InputType.药品 Then
'        Set RecReturn = Frm药品选择器.ShowME(Me, 1, 0, cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex))
        
        Set RecReturn = frmSelector.showMe(Me, 0, 1, , , , cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), 0, True, True, True, 0, False)

        If RecReturn.RecordCount = 0 Then
            Call zlControl.TxtSelAll(txtInput(Index))
            Exit Sub
        End If
            
        If gint药品名称显示 = 1 Then
            txtInput(Index).Text = "[" & RecReturn!药品编码 & "]" & IIf(IsNull(RecReturn!商品名), RecReturn!通用名, RecReturn!商品名)
        Else
            txtInput(Index).Text = "[" & RecReturn!药品编码 & "]" & RecReturn!通用名
        End If
        txtInput(Index).Tag = RecReturn!药品id
    Else
        If GetSeleterReturn(Index, txtInput(Index), "") = False Then
            Call zlControl.TxtSelAll(txtInput(Index))
        End If
    End If
End Sub


Private Sub Form_Load()
    Dim dateCurr As Date
    
    Set Me.Icon = Nothing
    dateCurr = Sys.Currentdate
    dtp开始时间.Value = CDate(Format(dateCurr, "YYYY-MM") & "-01 00:00:00")
    dtp结束时间.Value = dateCurr
    
    Call LoadStock
    Call IniGrid
    
    If mstrNo <> "" Then
        Call GetDate(mstrNo)
    Else
        txt填制人.Text = UserInfo.用户姓名
        Txt填制日期.Text = Format(dateCurr, "YYYY-MM-DD HH:MM:SS")
    End If
End Sub

Private Sub Form_Resize()
    If Me.Width < 13365 Then Me.Width = 13365
    If Me.Height < 8715 Then Me.Height = 8715
    
    With picBill
        .Top = 0
        .Left = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight - .Top - fraControl.Height - 100
    End With
    
    With fraControl
        .Top = Me.ScaleHeight - .Height - 100
        .Left = 0
        .Width = Me.ScaleWidth
        .Height = 615
    End With
    
    With fraCondition
        .Top = 50
        .Left = 100
        .Width = picBill.Width - 100
    End With
    
    With vsfBill
        .Top = fraCondition.Top + fraCondition.Height + 100
        .Left = fraCondition.Left
        .Width = fraCondition.Width
        .Height = picBill.Height - fraCondition.Top - fraCondition.Height - fraComment.Height - 100
    End With
    
    With fraComment
        .Top = picBill.Height - .Height
        .Left = 0
        .Width = picBill.Width
    End With
    
    With txt摘要
        .Width = fraComment.Width - .Top - 200
    End With
    
    RestoreWinState Me, App.ProductName, MStrCaption
End Sub


Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName, MStrCaption
    Call ReleaseSelectorRS
End Sub

Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strInput As String
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    strInput = Trim(UCase(txtFind.Text))
    If strInput = "" Then Exit Sub
    
    Call FindGridRow(strInput)
End Sub


Private Sub txtInput_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtInput(Index)
End Sub
Private Sub txtInput_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Trim(txtInput(Index).Text) = "" Then Exit Sub
    
    Dim sngLeft As Single
    Dim sngTop As Single
    Dim RecReturn As Recordset
    Dim strkey As String
    
    If Index = InputType.药品 Then
        If KeyCode <> vbKeyReturn Then Exit Sub
        If Trim(txtInput(Index).Text) = "" Then Exit Sub
        sngLeft = Me.Left + fraCondition.Left + txtInput(Index).Left
        sngTop = Me.Top + fraCondition.Top + txtInput(Index).Top + txtInput(Index).Height + Me.Height - Me.ScaleHeight '  50
        If sngTop + 3630 > Screen.Height Then
            sngTop = sngTop - txtInput(Index).Height - 3630
        End If
        
        strkey = Trim(txtInput(Index).Text)
        If Mid(strkey, 1, 1) = "[" Then
            If InStr(2, strkey, "]") <> 0 Then
                strkey = Mid(strkey, 2, InStr(2, strkey, "]") - 2)
            Else
                strkey = Mid(strkey, 2)
            End If
        End If
        
        If grsMaster.State = adStateClosed Then
            Call SetSelectorRS(1, "药品外购入库管理", cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex))
        End If
        
'        Set RecReturn = Frm药品多选选择器.ShowME(Me, 1, , cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), strkey, sngLeft, sngTop)
        Set RecReturn = frmSelector.showMe(Me, 1, 1, strkey, sngLeft, sngTop, cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), cboStock.ItemData(cboStock.ListIndex), 0, True, True, True, 0, False)
        
        If RecReturn.RecordCount = 0 Then
            Call zlControl.TxtSelAll(txtInput(Index))
            Exit Sub
        End If
        
        If gint药品名称显示 = 1 Then
            txtInput(Index).Text = "[" & RecReturn!药品编码 & "]" & IIf(IsNull(RecReturn!商品名), RecReturn!通用名, RecReturn!商品名)
        Else
            txtInput(Index).Text = "[" & RecReturn!药品编码 & "]" & RecReturn!通用名
        End If
        txtInput(Index).Tag = RecReturn!药品id
    Else
        If GetSeleterReturn(Index, txtInput(Index), Trim(txtInput(Index).Text)) = False Then
            Call zlControl.TxtSelAll(txtInput(Index))
        End If
    End If
End Sub


Private Sub txtInput_KeyPress(Index As Integer, KeyAscii As Integer)
    If InStr(" ';", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub vsfBill_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'    With vsfBill
'        Select Case Col
'            Case 明细列表.药品名称, 明细列表.供应商, 明细列表.生产商
'                .ColComboList(Col) = "..."
'        End Select
'    End With
End Sub

Private Sub vsfBill_AfterSort(ByVal Col As Long, Order As Integer)
    Call RefreshSerialNumber
End Sub

Private Sub vsfBill_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
'    With vsfBill
'        Select Case Col
'            Case 明细列表.药品名称
'                Call GetSeleterReturn(InputType.药品, vsfBill, "")
'            Case 明细列表.供应商
'                Call GetSeleterReturn(InputType.供应商, vsfBill, "")
'            Case 明细列表.生产商
'                Call GetSeleterReturn(InputType.生产商, vsfBill, "")
'        End Select
'    End With
End Sub
Private Sub vsfBill_EnterCell()
    With vsfBill
        .Editable = flexEDNone
        If .Row < 1 Then Exit Sub
        If .TextMatrix(.Row, 明细列表.药品id) = "" Then Exit Sub
        
        Select Case .Col
'        Case 明细列表.药品名称, 明细列表.供应商, 明细列表.生产商
'            .ColComboList(.Col) = "..."
'            .Editable = flexEDKbdMouse
        Case 明细列表.数量 ', 明细列表.成本价, 明细列表.批号
            .Editable = flexEDKbdMouse
        End Select
    End With
End Sub

Private Sub vsfBill_KeyDown(KeyCode As Integer, Shift As Integer)
    With vsfBill
        If KeyCode = vbKeyDelete Then
            If .Row < 1 Then Exit Sub
            If .TextMatrix(.Row, 明细列表.药品id) = "" Then Exit Sub
            
            If MsgBox("是否删除第" & .Row & "行的退药记录？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                .RemoveItem .Row
                Call RefreshSerialNumber
            End If
        End If
        
'        Select Case .Col
'        Case 明细列表.药品名称, 明细列表.供应商, 明细列表.生产商
'            If KeyCode <> vbKeyReturn Then
'                .ColComboList(.Col) = ""
'            End If
'        End Select
        
        If txtFind.Visible And KeyCode = vbKeyF3 Then
            Call txtFind_KeyDown(vbKeyReturn, 0)
        End If
    End With
End Sub


Private Sub vsfBill_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    With vsfBill
        If Trim(.EditText) = "" Then Exit Sub

'        Select Case Col
'            Case 明细列表.药品名称
'                Call GetSeleterReturn(InputType.药品, vsfBill, Trim(.EditText))
'            Case 明细列表.供应商
'                Call GetSeleterReturn(InputType.供应商, vsfBill, Trim(.EditText))
'            Case 明细列表.生产商
'                Call GetSeleterReturn(InputType.生产商, vsfBill, Trim(.EditText))
'        End Select
    End With
End Sub

Private Sub vsfBill_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If InStr(" ';", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    
    '只能输入数字
    If Col = 明细列表.数量 Then
        If InStr("1234567890" + Chr(46) + Chr(8) + Chr(13), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0
        End If
    End If
    
'    '只能输入数字，小数点
'    If Col = 明细列表.成本价 Then
'        If InStr(".1234567890" + Chr(46) + Chr(8) + Chr(13), Chr(KeyAscii)) = 0 Then
'            KeyAscii = 0
'        End If
'    End If
End Sub
Private Sub vsfBill_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfBill
        Select Case Col
        Case 明细列表.数量
            .TextMatrix(Row, 明细列表.数量) = Val(.EditText)
            .TextMatrix(Row, 明细列表.成本金额) = zlStr.FormatEx(Val(.TextMatrix(Row, 明细列表.数量)) * Val(.TextMatrix(Row, 明细列表.成本价)), 2, , True)
'        Case 明细列表.成本价
'            .EditText = zlStr.FormatEx(Val(.EditText), 5)
'            .TextMatrix(Row, 明细列表.成本价) = .EditText
'            .TextMatrix(Row, 明细列表.成本金额) = zlStr.FormatEx(Val(.TextMatrix(Row, 明细列表.数量)) * Val(.TextMatrix(Row, 明细列表.成本价)), 2)
        End Select
    End With
End Sub


