VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmLabItemOption 
   BorderStyle     =   0  'None
   Caption         =   "检验项目执行选项"
   ClientHeight    =   4410
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8025
   LinkTopic       =   "Form1"
   ScaleHeight     =   4410
   ScaleWidth      =   8025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame fraLine 
      Height          =   15
      Left            =   -30
      TabIndex        =   34
      Top             =   1785
      Width           =   7485
   End
   Begin VB.PictureBox picEdit 
      BorderStyle     =   0  'None
      Height          =   1995
      Left            =   3825
      ScaleHeight     =   1995
      ScaleWidth      =   4125
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1860
      Width           =   4125
      Begin VB.TextBox txt急诊耗时 
         Height          =   300
         Left            =   1155
         MaxLength       =   2
         TabIndex        =   45
         Top             =   630
         Width           =   615
      End
      Begin VB.TextBox txt送检时限 
         Height          =   300
         Left            =   1140
         MaxLength       =   2
         TabIndex        =   36
         Top             =   975
         Width           =   615
      End
      Begin VB.TextBox txt跟踪天数 
         Height          =   300
         Left            =   1155
         MaxLength       =   3
         TabIndex        =   20
         Top             =   0
         Width           =   615
      End
      Begin VB.TextBox txt耗时标准 
         Height          =   300
         Left            =   1155
         MaxLength       =   2
         TabIndex        =   22
         Top             =   315
         Width           =   615
      End
      Begin VB.ComboBox cbo耗时单位 
         Height          =   300
         Left            =   1845
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   315
         Width           =   825
      End
      Begin VB.TextBox txt报告地点 
         Height          =   300
         Left            =   645
         MaxLength       =   50
         TabIndex        =   25
         Top             =   1305
         Width           =   3000
      End
      Begin VB.TextBox txt报告说明 
         Height          =   300
         Left            =   645
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   27
         Top             =   1635
         Width           =   3000
      End
      Begin VB.Label lbl急诊耗时 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "急诊标本        分种后可取报告"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   375
         TabIndex        =   46
         Top             =   690
         Width           =   2700
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "送检标本超过        分钟后拒绝接收"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   30
         TabIndex        =   35
         Top             =   1035
         Width           =   3060
      End
      Begin VB.Label lbl跟踪天数 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "默认跟踪天数        ,按此对比历史结果"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   30
         TabIndex        =   19
         Top             =   45
         Width           =   3330
      End
      Begin VB.Label lbl执行时间 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "项目执行时间                  后可取报告"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   30
         TabIndex        =   21
         Top             =   360
         Width           =   3600
      End
      Begin VB.Label lbl报告地点 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "地点"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   240
         TabIndex        =   24
         Top             =   1365
         Width           =   360
      End
      Begin VB.Label lbl报告说明 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "说明"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   240
         TabIndex        =   26
         Top             =   1665
         Width           =   360
      End
   End
   Begin VB.Frame fraAppTo 
      Height          =   510
      Left            =   150
      TabIndex        =   28
      Top             =   3855
      Width           =   7710
      Begin VB.OptionButton optApplyTo 
         Caption         =   "所有检验项目"
         Height          =   180
         Index           =   2
         Left            =   6045
         TabIndex        =   32
         Top             =   195
         Width           =   1395
      End
      Begin VB.OptionButton optApplyTo 
         Caption         =   "所有""临检""类项目"
         Height          =   180
         Index           =   1
         Left            =   3360
         TabIndex        =   31
         Top             =   195
         Width           =   2670
      End
      Begin VB.OptionButton optApplyTo 
         Caption         =   "仅本项目"
         Height          =   180
         Index           =   0
         Left            =   2250
         TabIndex        =   30
         Top             =   195
         Value           =   -1  'True
         Width           =   1020
      End
      Begin VB.Label lblApplyTo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "保存时将特性同时应用于："
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   75
         TabIndex        =   29
         Top             =   195
         Width           =   2160
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vfg采集方式 
      Height          =   1530
      Left            =   120
      TabIndex        =   17
      Top             =   2085
      Width           =   3615
      _cx             =   6376
      _cy             =   2699
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
      BackColorFixed  =   15790320
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16772055
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   -1  'True
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
   Begin VB.PictureBox picDept 
      BorderStyle     =   0  'None
      Height          =   1650
      Left            =   105
      ScaleHeight     =   1650
      ScaleWidth      =   7875
      TabIndex        =   33
      Top             =   45
      Width           =   7875
      Begin VB.CheckBox chk服务对象 
         Caption         =   "仅用于体检病人(&I)"
         Height          =   225
         Index           =   2
         Left            =   5280
         TabIndex        =   41
         Top             =   0
         Width           =   1890
      End
      Begin VB.ComboBox cbo执行科室 
         Height          =   300
         Index           =   2
         ItemData        =   "frmLabItemOption.frx":0000
         Left            =   6045
         List            =   "frmLabItemOption.frx":0002
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   660
         Width           =   1800
      End
      Begin VB.ComboBox cbo诊疗单据 
         Height          =   300
         Index           =   2
         ItemData        =   "frmLabItemOption.frx":0004
         Left            =   6045
         List            =   "frmLabItemOption.frx":0006
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   285
         Width           =   1800
      End
      Begin VB.ComboBox cbo默认仪器 
         Height          =   300
         Index           =   2
         ItemData        =   "frmLabItemOption.frx":0008
         Left            =   6045
         List            =   "frmLabItemOption.frx":000A
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   1035
         Width           =   1800
      End
      Begin VB.CheckBox chk仪器分解 
         Caption         =   "按组成指标默认仪器分配"
         Height          =   195
         Index           =   2
         Left            =   5535
         TabIndex        =   37
         Top             =   1425
         Width           =   2475
      End
      Begin VB.CheckBox chk仪器分解 
         Caption         =   "按组成指标默认仪器分配"
         Height          =   195
         Index           =   1
         Left            =   2880
         TabIndex        =   15
         Top             =   1425
         Width           =   2295
      End
      Begin VB.CheckBox chk仪器分解 
         Caption         =   "按组成指标默认仪器分配"
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   7
         Top             =   1425
         Width           =   2550
      End
      Begin VB.ComboBox cbo默认仪器 
         Height          =   300
         Index           =   1
         ItemData        =   "frmLabItemOption.frx":000C
         Left            =   3390
         List            =   "frmLabItemOption.frx":000E
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1035
         Width           =   1800
      End
      Begin VB.ComboBox cbo默认仪器 
         Height          =   300
         Index           =   0
         ItemData        =   "frmLabItemOption.frx":0010
         Left            =   780
         List            =   "frmLabItemOption.frx":0012
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1035
         Width           =   1800
      End
      Begin VB.ComboBox cbo诊疗单据 
         Height          =   300
         Index           =   1
         ItemData        =   "frmLabItemOption.frx":0014
         Left            =   3390
         List            =   "frmLabItemOption.frx":0016
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   285
         Width           =   1800
      End
      Begin VB.ComboBox cbo执行科室 
         Height          =   300
         Index           =   1
         ItemData        =   "frmLabItemOption.frx":0018
         Left            =   3390
         List            =   "frmLabItemOption.frx":001F
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   660
         Width           =   1800
      End
      Begin VB.ComboBox cbo诊疗单据 
         Height          =   300
         Index           =   0
         ItemData        =   "frmLabItemOption.frx":0030
         Left            =   795
         List            =   "frmLabItemOption.frx":0037
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   285
         Width           =   1800
      End
      Begin VB.ComboBox cbo执行科室 
         Height          =   300
         Index           =   0
         ItemData        =   "frmLabItemOption.frx":0048
         Left            =   780
         List            =   "frmLabItemOption.frx":004F
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   660
         Width           =   1800
      End
      Begin VB.CheckBox chk服务对象 
         Caption         =   "可用于门诊病人(&O)"
         Height          =   225
         Index           =   0
         Left            =   0
         TabIndex        =   0
         Top             =   0
         Value           =   1  'Checked
         Width           =   1950
      End
      Begin VB.CheckBox chk服务对象 
         Caption         =   "可用于住院病人(&I)"
         Height          =   225
         Index           =   1
         Left            =   2625
         TabIndex        =   8
         Top             =   0
         Value           =   1  'Checked
         Width           =   1890
      End
      Begin VB.Label lbl诊疗单据 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "诊疗单据"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   2
         Left            =   5295
         TabIndex        =   43
         Top             =   345
         Width           =   705
      End
      Begin VB.Label lbl执行科室 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "执行科室"
         Height          =   180
         Index           =   2
         Left            =   5295
         TabIndex        =   44
         Top             =   720
         Width           =   720
      End
      Begin VB.Label lbl默认仪器 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "默认仪器"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   2
         Left            =   5295
         TabIndex        =   42
         Top             =   1095
         Width           =   705
      End
      Begin VB.Label lbl默认仪器 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "默认仪器"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   2610
         TabIndex        =   13
         Top             =   1095
         Width           =   735
      End
      Begin VB.Label lbl默认仪器 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "默认仪器"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   30
         TabIndex        =   5
         Top             =   1095
         Width           =   720
      End
      Begin VB.Label lbl诊疗单据 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "诊疗单据"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   2610
         TabIndex        =   9
         Top             =   345
         Width           =   750
      End
      Begin VB.Label lbl诊疗单据 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "诊疗单据"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   30
         TabIndex        =   1
         Top             =   345
         Width           =   720
      End
      Begin VB.Label lbl执行科室 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "执行科室"
         Height          =   180
         Index           =   0
         Left            =   30
         TabIndex        =   3
         Top             =   720
         Width           =   720
      End
      Begin VB.Label lbl执行科室 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "执行科室"
         Height          =   180
         Index           =   1
         Left            =   2610
         TabIndex        =   11
         Top             =   720
         Width           =   780
      End
   End
   Begin VB.Label lbl采集方式 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "标本采集方式(&G)"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   105
      TabIndex        =   16
      Top             =   1845
      Width           =   1350
   End
End
Attribute VB_Name = "frmLabItemOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngItemID As Long          '当前显示的项目id
Private mint组合 As Integer

Private Enum mCol
    ID = 0: 标志: 编码: 名称: 撤档时间
End Enum

'临时变量
Dim lngCount As Long
Dim strTemp As String, aryTemp() As String

'--------------------------------------------
'以下为窗体公共方法
'--------------------------------------------
Public Function zlRefresh(lngItemId As Long) As Boolean
    '功能：根据项目id刷新当前显示内容
    '参数：当前项目id
    Dim rsTemp As New ADODB.Recordset
    Dim j As Integer
    mlngItemID = lngItemId: mint组合 = 0
    
    '清除此前项目的显示
    Me.chk服务对象(0).Value = 0: Me.cbo诊疗单据(0).ListIndex = -1: Me.cbo执行科室(0).ListIndex = -1
    Me.cbo默认仪器(0).ListIndex = -1: Me.chk仪器分解(0).Value = vbUnchecked: Me.chk仪器分解(0).Enabled = False
    Me.chk服务对象(1).Value = 0: Me.cbo诊疗单据(1).ListIndex = -1: Me.cbo执行科室(1).ListIndex = -1
    Me.cbo默认仪器(1).ListIndex = -1: Me.chk仪器分解(1).Value = vbUnchecked: Me.chk仪器分解(1).Enabled = False
    Me.chk服务对象(2).Value = 0: Me.cbo诊疗单据(2).ListIndex = -1: Me.cbo执行科室(2).ListIndex = -1
    Me.cbo默认仪器(2).ListIndex = -1: Me.chk仪器分解(2).Value = vbUnchecked: Me.chk仪器分解(2).Enabled = False
    Me.txt跟踪天数.Text = "": Me.txt耗时标准.Text = "": Me.cbo耗时单位.ListIndex = 0
    Me.txt报告地点.Text = "": Me.txt报告说明.Text = "": Me.txt送检时限.Text = ""
    Me.txt急诊耗时.Text = ""
    
    With Me.vfg采集方式
        For lngCount = .FixedRows To .Rows - 1
            .Row = lngCount: .Col = mCol.编码
            .CellChecked = flexUnchecked
        Next
    End With
    If lngItemId = 0 Then zlRefresh = True: Exit Function
    
    '获取指定项目的信息
    Err = 0: On Error GoTo ErrHand
    '装入相同检验类型的仪器
    gstrSql = "Select M.ID, M.编码, M.名称 From 检验仪器 M, 诊疗项目目录 I Where M.仪器类型 = I.操作类型 And I.ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemId)
    With rsTemp
        Me.cbo默认仪器(0).Clear: Me.cbo默认仪器(1).Clear: Me.cbo默认仪器(2).Clear
        Do While Not .EOF
            Me.cbo默认仪器(0).AddItem !编码 & "-" & !名称: Me.cbo默认仪器(0).ItemData(Me.cbo默认仪器(0).NewIndex) = !ID
            Me.cbo默认仪器(1).AddItem !编码 & "-" & !名称: Me.cbo默认仪器(1).ItemData(Me.cbo默认仪器(1).NewIndex) = !ID
            Me.cbo默认仪器(2).AddItem !编码 & "-" & !名称: Me.cbo默认仪器(2).ItemData(Me.cbo默认仪器(2).NewIndex) = !ID
            .MoveNext
        Loop
    End With
    
    gstrSql = "Select I.服务对象, I.操作类型, I.组合项目, O.门诊仪器id, O.门诊仪器分解, O.住院仪器id, O.住院仪器分解, O.跟踪天数," & vbNewLine & _
            "       O.耗时标准, O.耗时单位, O.取报告地点, O.附加说明,O.送检时限,O.体检仪器id,O.体检仪器分解,O.急诊耗时 " & vbNewLine & _
            "From 诊疗项目目录 I, 检验项目选项 O" & vbNewLine & _
            "Where I.ID = O.诊疗项目id(+) And I.ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemId)
    With rsTemp
        If .RecordCount > 0 Then
            If Val("" & !组合项目) = 1 Then mint组合 = 1: Me.chk仪器分解(0).Enabled = True: Me.chk仪器分解(1).Enabled = True
            
            If Val("" & !服务对象) = 4 Then
                Me.chk服务对象(2).Value = vbChecked: Me.chk服务对象(0).Value = vbUnchecked: Me.chk服务对象(1).Value = vbUnchecked
            ElseIf Val("" & !服务对象) = 3 Then
                Me.chk服务对象(0).Value = vbChecked: Me.chk服务对象(1).Value = vbChecked
            ElseIf Val("" & !服务对象) = 1 Then
                Me.chk服务对象(0).Value = vbChecked
            ElseIf Val("" & !服务对象) = 2 Then
                Me.chk服务对象(1).Value = vbChecked
            End If
            Me.optApplyTo(1).Caption = "所有""" & !操作类型 & """类项目"
            
            For lngCount = 0 To Me.cbo默认仪器(0).ListCount - 1
                If Me.cbo默认仪器(0).ItemData(lngCount) = Val("" & !门诊仪器id) Then Me.cbo默认仪器(0).ListIndex = lngCount: Exit For
            Next
            Me.chk仪器分解(0).Value = IIf(Val("" & !门诊仪器分解) = 0, vbUnchecked, vbChecked)
            
            For lngCount = 0 To Me.cbo默认仪器(1).ListCount - 1
                If Me.cbo默认仪器(1).ItemData(lngCount) = Val("" & !住院仪器id) Then Me.cbo默认仪器(1).ListIndex = lngCount: Exit For
            Next
            Me.chk仪器分解(1).Value = IIf(Val("" & !住院仪器分解) = 0, vbUnchecked, vbChecked)
            
            For lngCount = 0 To Me.cbo默认仪器(2).ListCount - 1
                If Me.cbo默认仪器(2).ItemData(lngCount) = Val("" & !体检仪器id) Then Me.cbo默认仪器(2).ListIndex = lngCount: Exit For
            Next
            Me.chk仪器分解(2).Value = IIf(Val("" & !体检仪器分解) = 0, vbUnchecked, vbChecked)
            
            Me.txt跟踪天数.Text = "" & !跟踪天数
            Me.txt耗时标准.Text = "" & !耗时标准
            For lngCount = 0 To Me.cbo耗时单位.ListCount - 1
                If Me.cbo耗时单位.List(lngCount) = !耗时单位 Then Me.cbo耗时单位.ListIndex = lngCount: Exit For
            Next
            Me.txt急诊耗时.Text = "" & !急诊耗时
            
            Me.txt报告地点.Text = "" & !取报告地点
            Me.txt报告说明.Text = "" & !附加说明
            Me.txt送检时限.Text = IIf(Val(Nvl(!送检时限)) = 0, "", Nvl(!送检时限))
        End If
    End With
    
    gstrSql = "Select 用法id From 诊疗用法用量 Where 项目id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemId)
    Do While Not rsTemp.EOF
        With Me.vfg采集方式
            For lngCount = .FixedRows To .Rows - 1
                .Row = lngCount: .Col = mCol.编码
                If Val(.TextMatrix(lngCount, mCol.撤档时间)) = 0 Then
                    If Val(.TextMatrix(lngCount, mCol.ID)) = Val("" & rsTemp!用法ID) Then .CellChecked = flexChecked
                ElseIf Format(.TextMatrix(lngCount, mCol.撤档时间), "yyyy-mm-dd") = "3000-01-01" Then
                    If Val(.TextMatrix(lngCount, mCol.ID)) = Val("" & rsTemp!用法ID) Then .CellChecked = flexChecked
                Else
                    If Val(.TextMatrix(lngCount, mCol.ID)) = Val("" & rsTemp!用法ID) Then .CellChecked = flexTSGrayed
                End If
            Next
        End With
        rsTemp.MoveNext
    Loop
    
    gstrSql = "Select 病人来源, 执行科室id From 诊疗执行科室 Where 诊疗项目id = [1] And 开单科室id Is Null"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemId)
    Do While Not rsTemp.EOF
        If Val("" & rsTemp!病人来源) = 4 Then
            For lngCount = 0 To Me.cbo执行科室(2).ListCount - 1
                If Me.cbo执行科室(2).ItemData(lngCount) = rsTemp!执行科室ID Then Me.cbo执行科室(2).ListIndex = lngCount: Exit For
            Next
        ElseIf Val("" & rsTemp!病人来源) = 1 Then
            For lngCount = 0 To Me.cbo执行科室(0).ListCount - 1
                If Me.cbo执行科室(0).ItemData(lngCount) = rsTemp!执行科室ID Then Me.cbo执行科室(0).ListIndex = lngCount: Exit For
            Next
        ElseIf Val("" & rsTemp!病人来源) = 2 Then
            For lngCount = 0 To Me.cbo执行科室(1).ListCount - 1
                If Me.cbo执行科室(1).ItemData(lngCount) = rsTemp!执行科室ID Then Me.cbo执行科室(1).ListIndex = lngCount: Exit For
            Next
        End If
        rsTemp.MoveNext
    Loop
    
    gstrSql = "Select 应用场合, 病历文件id From 病历单据应用 Where 诊疗项目id = [1] And 应用场合 In (1, 2, 4)"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemId)
    Do While Not rsTemp.EOF
        If Val("" & rsTemp!应用场合) = 1 Then
            For lngCount = 0 To Me.cbo诊疗单据(0).ListCount - 1
                If Me.cbo诊疗单据(0).ItemData(lngCount) = rsTemp!病历文件id Then Me.cbo诊疗单据(0).ListIndex = lngCount: Exit For
            Next
        ElseIf Val("" & rsTemp!应用场合) = 2 Then
            For lngCount = 0 To Me.cbo诊疗单据(1).ListCount - 1
                If Me.cbo诊疗单据(1).ItemData(lngCount) = rsTemp!病历文件id Then Me.cbo诊疗单据(1).ListIndex = lngCount: Exit For
            Next
        ElseIf Val("" & rsTemp!应用场合) = 4 Then
            For lngCount = 0 To Me.cbo诊疗单据(2).ListCount - 1
                If Me.cbo诊疗单据(2).ItemData(lngCount) = rsTemp!病历文件id Then Me.cbo诊疗单据(2).ListIndex = lngCount: Exit For
            Next
        End If
        rsTemp.MoveNext
    Loop
    
    zlRefresh = True: Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlRefresh = False: Exit Function
End Function

Public Function zlEditStart() As Boolean
    '功能：开始项目编辑
    '参数： lngItemId-指定编辑的项目
    Me.Tag = "编辑": Call Form_Resize
    zlEditStart = True: Exit Function
End Function

Public Sub zlEditCancel()
    '功能：放弃正在进行的编辑
    Me.Tag = "": Call Form_Resize
    Call Me.zlRefresh(mlngItemID)
End Sub

Public Function zlEditSave() As Long
    '功能：保存正在进行的编辑,并返回正在编辑项目id,保存失败返回0
    Dim strList As String
    Dim strDept0 As String, strDept1 As String, strDept2 As String
    Dim strBill0 As String, strBill1 As String, strBill2 As String
    Dim strApt0 As String, strApt1 As String, strApt2 As String
    Dim strAllot0 As String, strAllot1 As String
    
    If Me.cbo耗时单位.ListIndex = -1 Then Me.cbo耗时单位.ListIndex = 0
    '数据保存语句组织
    strList = ""
    With Me.vfg采集方式
        For lngCount = .FixedRows To .Rows - 1
            .Row = lngCount: .Col = mCol.编码
            If .CellChecked = flexChecked Then strList = strList & "," & .TextMatrix(lngCount, mCol.ID)
            If .CellChecked = flexTSGrayed Then strList = strList & "," & .TextMatrix(lngCount, mCol.ID)
        Next
    End With
    If strList <> "" Then strList = Mid(strList, 2)

    If Me.cbo执行科室(0).ListIndex = -1 Then
        strDept0 = "Null"
    Else
        strDept0 = Me.cbo执行科室(0).ItemData(Me.cbo执行科室(0).ListIndex)
    End If
    If Me.cbo执行科室(1).ListIndex = -1 Then
        strDept1 = "Null"
    Else
        strDept1 = Me.cbo执行科室(1).ItemData(Me.cbo执行科室(1).ListIndex)
    End If
    If Me.cbo诊疗单据(0).ListIndex = -1 Then
        strBill0 = "Null"
    Else
        strBill0 = Me.cbo诊疗单据(0).ItemData(Me.cbo诊疗单据(0).ListIndex)
    End If
    If Me.cbo诊疗单据(1).ListIndex = -1 Then
        strBill1 = "Null"
    Else
        strBill1 = Me.cbo诊疗单据(1).ItemData(Me.cbo诊疗单据(1).ListIndex)
    End If
    If Me.cbo默认仪器(0).ListIndex = -1 Then
        strApt0 = "Null"
    Else
        strApt0 = Me.cbo默认仪器(0).ItemData(Me.cbo默认仪器(0).ListIndex)
    End If
    If Me.cbo默认仪器(1).ListIndex = -1 Then
        strApt1 = "Null"
    Else
        strApt1 = Me.cbo默认仪器(1).ItemData(Me.cbo默认仪器(1).ListIndex)
    End If
    strAllot0 = IIf(Me.chk仪器分解(0).Value = vbChecked, 1, 0)
    strAllot1 = IIf(Me.chk仪器分解(1).Value = vbChecked, 1, 0)
    
    If Me.chk服务对象(2) = vbChecked Then
        gstrSql = mlngItemID & ",'" & strList & "',4," & strDept0 & "," & strDept1 & "," & strBill0 & "," & strBill1
        gstrSql = gstrSql & "," & strApt0 & "," & strAllot0 & "," & strApt1 & "," & strAllot1
    ElseIf Me.chk服务对象(0).Value = vbChecked And Me.chk服务对象(1).Value = vbChecked Then
        gstrSql = mlngItemID & ",'" & strList & "',3," & strDept0 & "," & strDept1 & "," & strBill0 & "," & strBill1
        gstrSql = gstrSql & "," & strApt0 & "," & strAllot0 & "," & strApt1 & "," & strAllot1
    ElseIf Me.chk服务对象(0).Value <> vbChecked And Me.chk服务对象(1).Value = vbChecked Then
        gstrSql = mlngItemID & ",'" & strList & "',2,Null," & strDept1 & ",Null," & strBill1
        gstrSql = gstrSql & ",Null,0," & strApt1 & "," & strAllot1
    ElseIf Me.chk服务对象(0).Value = vbChecked And Me.chk服务对象(1).Value <> vbChecked Then
        gstrSql = mlngItemID & ",'" & strList & "',1," & strDept0 & ",Null," & strBill0 & ",Null"
        gstrSql = gstrSql & "," & strApt0 & "," & strAllot0 & ",Null,0"
    Else
        gstrSql = mlngItemID & ",'" & strList & "',0,Null,Null,Null,Null,Null,Null,Null,Null"
    End If
    gstrSql = gstrSql & "," & Val(Me.txt跟踪天数.Text)
    gstrSql = gstrSql & "," & Val(Me.txt耗时标准.Text)
    gstrSql = gstrSql & ",'" & Me.cbo耗时单位.Text & "'"
    gstrSql = gstrSql & ",'" & Me.txt报告地点.Text & "'"
    gstrSql = gstrSql & ",'" & Me.txt报告说明.Text & "'"
    
    If optApplyTo.Item(0).Value = True Then
        gstrSql = gstrSql & "," & 0
    ElseIf optApplyTo.Item(1).Value = True Then
        gstrSql = gstrSql & "," & 1
    ElseIf optApplyTo.Item(2).Value = True Then
        gstrSql = gstrSql & "," & 2
    End If
    gstrSql = gstrSql & "," & IIf(Trim(Me.txt送检时限) = "", "NULL", Val(Me.txt送检时限))
    
    strDept2 = "Null"
    strBill2 = "Null"
    strApt2 = "Null"
    If Me.cbo执行科室(2).ListIndex <> -1 Then strDept2 = Me.cbo执行科室(2).ItemData(Me.cbo执行科室(2).ListIndex)
    If Me.cbo诊疗单据(2).ListIndex <> -1 Then strBill2 = Me.cbo诊疗单据(2).ItemData(Me.cbo诊疗单据(2).ListIndex)
    If Me.cbo默认仪器(2).ListIndex <> -1 Then strApt2 = Me.cbo默认仪器(2).ItemData(Me.cbo默认仪器(2).ListIndex)
    gstrSql = gstrSql & "," & strDept2 & "," & strBill2 & "," & strApt2 & "," & IIf(Me.chk仪器分解(2).Value = vbChecked, 1, 0)
    
    gstrSql = gstrSql & "," & IIf(Trim(Me.txt急诊耗时) = "", "Null", Val(Me.txt急诊耗时))
    
    gstrSql = "Zl_检验项目选项_Edit(" & gstrSql & ")"
    Err = 0: On Error GoTo ErrHand
    Call SQLTest(App.ProductName, Me.Caption, gstrSql): gcnOracle.Execute gstrSql, , adCmdStoredProc: Call SQLTest
    
    Me.Tag = "": Call Form_Resize
    zlEditSave = mlngItemID: Exit Function
    
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlEditSave = 0: Exit Function
End Function



'--------------------------------------------
'以下为窗体控件响应事件
'--------------------------------------------
Private Sub cbo耗时单位_Click()
    If Me.cbo耗时单位.ListCount > 0 Then
        Me.lbl急诊耗时.Caption = "急诊标本        " & Me.cbo耗时单位.List(Me.cbo耗时单位.ListIndex) & "后可取报告"
    End If
End Sub

Private Sub cbo耗时单位_GotFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub cbo耗时单位_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo默认仪器_GotFocus(Index As Integer)
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub cbo默认仪器_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo诊疗单据_GotFocus(Index As Integer)
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub cbo诊疗单据_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo执行科室_GotFocus(Index As Integer)
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub cbo执行科室_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chk服务对象_Click(Index As Integer)


    
    If Me.chk服务对象(Index).Value Then
        Me.cbo执行科室(Index).Enabled = True: Me.cbo诊疗单据(Index).Enabled = True
        Me.cbo默认仪器(Index).Enabled = True: Me.chk仪器分解(Index).Enabled = (mint组合 = 1)
        
        If Index = 2 Then
            If Me.chk服务对象(0).Value <> 0 Then
                Me.chk服务对象(0).Value = 0
                Me.cbo执行科室(0).Enabled = False: Me.cbo诊疗单据(0).Enabled = False
                Me.cbo默认仪器(0).Enabled = False: Me.chk仪器分解(0).Enabled = False
            End If
            If Me.chk服务对象(1).Value <> 0 Then
                Me.chk服务对象(1).Value = 0
                Me.cbo执行科室(1).Enabled = False: Me.cbo诊疗单据(1).Enabled = False
                Me.cbo默认仪器(1).Enabled = False: Me.chk仪器分解(1).Enabled = False
            End If
        Else
            If Me.chk服务对象(2).Value <> 0 Then
                Me.chk服务对象(2).Value = 0
                Me.cbo执行科室(2).Enabled = False: Me.cbo诊疗单据(2).Enabled = False
                Me.cbo默认仪器(2).Enabled = False: Me.chk仪器分解(2).Enabled = False
            End If
        End If
        
    Else
        Me.cbo执行科室(Index).Enabled = False: Me.cbo诊疗单据(Index).Enabled = False
        Me.cbo默认仪器(Index).Enabled = False: Me.chk仪器分解(Index).Enabled = False
    End If
    Me.cbo执行科室(2).Enabled = True: Me.cbo诊疗单据(2).Enabled = True
    
End Sub

Private Sub chk服务对象_GotFocus(Index As Integer)
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub chk服务对象_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chk仪器分解_GotFocus(Index As Integer)
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub chk仪器分解_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_Load()
    '基本数据装入
    Dim rsTemp As New ADODB.Recordset
    Dim i As Integer
    Dim j As Integer
    Err = 0: On Error GoTo ErrHand
    gstrSql = "Select I.ID, 0 As 标志, I.编码, I.名称,i.撤档时间 From 诊疗项目目录 I Where I.类别 = 'E' And I.操作类型 = '6' " & vbNewLine & _
              "  Order By I.编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    With Me.vfg采集方式
        .Redraw = flexRDNone
         Set .DataSource = rsTemp
        .TextMatrix(0, mCol.ID) = "ID": .TextMatrix(0, mCol.标志) = ""
        .TextMatrix(0, mCol.编码) = "编码": .TextMatrix(0, mCol.名称) = "名称"
        .TextMatrix(0, mCol.撤档时间) = "撤档时间"
        .ColWidth(mCol.ID) = 0: .ColWidth(mCol.标志) = 0: .ColWidth(mCol.撤档时间) = 0
        .ColWidth(mCol.编码) = 800: .ColWidth(mCol.名称) = 2000
        For j = .FixedRows To .Rows - 1
            If Val(.TextMatrix(j, mCol.撤档时间)) = 0 Or Format(.TextMatrix(j, mCol.撤档时间), "yyyy-mm-dd") = "3000-01-01" Then
                
            Else
                For i = mCol.编码 To .Cols - 1
                    .Cell(flexcpForeColor, j, i, j, i) = vbRed
                
                    .TextMatrix(j, i) = .TextMatrix(j, i) & "(已停用)"
                Next
            End If
        Next
        For lngCount = .FixedCols To .Cols - 1
            .FixedAlignment(lngCount) = flexAlignCenterCenter
            If .ColWidth(lngCount) = 0 Then .ColHidden(lngCount) = True
        Next
        .Redraw = flexRDDirect
    End With
    
    aryTemp = Split("分钟;小时;天", ";")
    Me.cbo耗时单位.Clear
    For lngCount = LBound(aryTemp) To UBound(aryTemp)
        Me.cbo耗时单位.AddItem aryTemp(lngCount)
    Next
    If Me.cbo耗时单位.ListCount > 0 Then Me.cbo耗时单位.ListIndex = 0
    
    cbo执行科室(0).Clear: cbo执行科室(1).Clear: cbo执行科室(2).Clear
    cbo诊疗单据(0).Clear: cbo诊疗单据(1).Clear: cbo诊疗单据(2).Clear
    Call zlControl.CboSetWidth(cbo诊疗单据(0).hWnd, 2650)
    Call zlControl.CboSetWidth(cbo诊疗单据(1).hWnd, 2650)
    Call zlControl.CboSetWidth(cbo诊疗单据(2).hWnd, 2650)
    
    '--- 门诊 体检
    gstrSql = "Select ID, 编码, 名称" & vbNewLine & _
            "From 部门表 D, 部门性质说明 P" & vbNewLine & _
            "Where D.ID = P.部门id And P.工作性质 = '检验' And P.服务对象 In (1, 3) And" & vbNewLine & _
            "      (To_Char(D.撤档时间, 'YYYY-MM-DD') = '3000-01-01' Or D.撤档时间 Is Null)"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    With rsTemp
        Do While Not .EOF
            Me.cbo执行科室(0).AddItem !编码 & "-" & !名称
            Me.cbo执行科室(0).ItemData(Me.cbo执行科室(0).NewIndex) = !ID
            
            Me.cbo执行科室(2).AddItem !编码 & "-" & !名称
            Me.cbo执行科室(2).ItemData(Me.cbo执行科室(2).NewIndex) = !ID
            
            .MoveNext
        Loop
    End With
    gstrSql = "Select ID, 编号, 名称 From 病历文件列表 Where 种类 = 7"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    With rsTemp
        Do While Not .EOF
            Me.cbo诊疗单据(0).AddItem !编号 & "-" & !名称
            Me.cbo诊疗单据(0).ItemData(Me.cbo诊疗单据(0).NewIndex) = !ID
            
            Me.cbo诊疗单据(2).AddItem !编号 & "-" & !名称
            Me.cbo诊疗单据(2).ItemData(Me.cbo诊疗单据(2).NewIndex) = !ID
            .MoveNext
        Loop
    End With
    '--- 住院
    gstrSql = "Select ID, 编码, 名称" & vbNewLine & _
            "From 部门表 D, 部门性质说明 P" & vbNewLine & _
            "Where D.ID = P.部门id And P.工作性质 = '检验' And P.服务对象 In (2, 3) And" & vbNewLine & _
            "      (To_Char(D.撤档时间, 'YYYY-MM-DD') = '3000-01-01' Or D.撤档时间 Is Null)"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    With rsTemp
        Do While Not .EOF
            Me.cbo执行科室(1).AddItem !编码 & "-" & !名称
            Me.cbo执行科室(1).ItemData(Me.cbo执行科室(1).NewIndex) = !ID
            .MoveNext
        Loop
    End With
    gstrSql = "Select ID, 编号, 名称 From 病历文件列表 Where 种类 = 7"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    With rsTemp
        Do While Not .EOF
            Me.cbo诊疗单据(1).AddItem !编号 & "-" & !名称
            Me.cbo诊疗单据(1).ItemData(Me.cbo诊疗单据(1).NewIndex) = !ID
            .MoveNext
        Loop
    End With
        
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    Me.fraAppTo.Top = Me.ScaleHeight - Me.fraAppTo.Height - 180
    If Me.Tag = "编辑" Then
        Me.vfg采集方式.Height = Me.fraAppTo.Top - Me.vfg采集方式.Top
        Me.picEdit.Height = Me.fraAppTo.Top - Me.picEdit.Top
        Me.picEdit.Enabled = True: Me.picDept.Enabled = True
        Me.fraAppTo.Enabled = True: Me.fraAppTo.Visible = True
    Else
        Me.vfg采集方式.Height = Me.ScaleHeight - Me.vfg采集方式.Top - 180
        Me.picEdit.Height = Me.ScaleHeight - Me.picEdit.Top - 180
        Me.picEdit.Enabled = False: Me.picDept.Enabled = False
        Me.fraAppTo.Enabled = False: Me.fraAppTo.Visible = False
    End If
End Sub

Private Sub optApplyTo_GotFocus(Index As Integer)
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub optApplyTo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub picEdit_Resize()
    Err = 0: On Error Resume Next
    Me.txt报告说明.Height = Me.picEdit.ScaleHeight - Me.txt报告说明.Top
End Sub

Private Sub txt报告地点_GotFocus()
    Me.txt报告地点.SelStart = 0: Me.txt报告地点.SelLength = 1000
End Sub

Private Sub txt报告地点_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt报告说明_GotFocus()
    Me.txt报告说明.SelStart = 0: Me.txt报告说明.SelLength = 1000
End Sub

Private Sub txt报告说明_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt跟踪天数_GotFocus()
    Me.txt跟踪天数.SelStart = 0: Me.txt跟踪天数.SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt跟踪天数_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt耗时标准_GotFocus()
    Me.txt耗时标准.SelStart = 0: Me.txt耗时标准.SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt耗时标准_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt送检时限_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt急诊耗时_GotFocus()
    Me.txt急诊耗时.SelStart = 0: Me.txt急诊耗时.SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt急诊耗时_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub vfg采集方式_DblClick()
    If Me.vfg采集方式.MouseRow < Me.vfg采集方式.FixedRows Then Exit Sub
    If Me.Tag <> "编辑" Then Exit Sub
    With Me.vfg采集方式
        If .Row < .FixedRows And .Row > .Rows - 1 Then Exit Sub
        .Col = mCol.编码
        If Val(.TextMatrix(.Row, mCol.撤档时间)) = 0 Or Format(.TextMatrix(.Row, mCol.撤档时间), "yyyy-mm-dd") = "3000-01-01" Then
                
        Else
            MsgBox "已停用的采集方式，不能勾选！", vbInformation, Me.Caption
            Exit Sub
        End If
        If .CellChecked = flexChecked Then
            .CellChecked = flexUnchecked
        Else
            .CellChecked = flexChecked
        End If
    End With
End Sub

Private Sub vfg采集方式_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If KeyAscii <> vbKeySpace Then Exit Sub
    Call vfg采集方式_DblClick
End Sub
