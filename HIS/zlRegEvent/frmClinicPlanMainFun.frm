VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmClinicPlanMainFun 
   BorderStyle     =   0  'None
   Caption         =   "功能菜单"
   ClientHeight    =   7740
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7740
   ScaleWidth      =   11895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picPlanBack 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   4965
      Left            =   7590
      ScaleHeight     =   4965
      ScaleWidth      =   3345
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   690
      Width           =   3345
      Begin VB.Frame frmMoveSplitY 
         Height          =   25
         Left            =   -150
         MousePointer    =   7  'Size N S
         TabIndex        =   5
         Top             =   2670
         Width           =   3735
      End
      Begin MSComctlLib.TreeView tvwPlan 
         Height          =   1425
         Left            =   390
         TabIndex        =   6
         Top             =   3270
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   2514
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   88
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         ImageList       =   "imgPlan16"
         Appearance      =   0
         OLEDragMode     =   1
      End
      Begin MSComctlLib.TreeView tvwPlanTemplet 
         Height          =   1815
         Left            =   240
         TabIndex        =   7
         Top             =   810
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   3201
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   88
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         ImageList       =   "imgPlan16"
         Appearance      =   0
         OLEDragMode     =   1
      End
      Begin VB.Image imgYearSelect 
         Height          =   120
         Left            =   2250
         Picture         =   "frmClinicPlanMainFun.frx":0000
         Top             =   2880
         Width           =   120
      End
      Begin VB.Label lblYearSelect 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2016年"
         Height          =   180
         Left            =   1680
         TabIndex        =   10
         Top             =   2850
         Width           =   540
      End
      Begin XtremeSuiteControls.ShortcutCaption sccPlan 
         Height          =   360
         Left            =   0
         TabIndex        =   9
         Top             =   2760
         Width           =   3165
         _Version        =   589884
         _ExtentX        =   5583
         _ExtentY        =   635
         _StockProps     =   6
         Caption         =   "出诊安排"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
      Begin XtremeSuiteControls.ShortcutCaption sccPlanTemplet 
         Height          =   360
         Left            =   0
         TabIndex        =   8
         Top             =   390
         Width           =   2505
         _Version        =   589884
         _ExtentX        =   4419
         _ExtentY        =   635
         _StockProps     =   6
         Caption         =   "出诊模板"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
   End
   Begin VB.PictureBox picBaseSetBack 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   4905
      Left            =   270
      ScaleHeight     =   4905
      ScaleWidth      =   2985
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   570
      Width           =   2985
      Begin XtremeSuiteControls.TaskPanel tplFunBase 
         Height          =   3045
         Left            =   540
         TabIndex        =   11
         Top             =   1110
         Width           =   1815
         _Version        =   589884
         _ExtentX        =   3201
         _ExtentY        =   5371
         _StockProps     =   64
         ItemLayout      =   2
         HotTrackStyle   =   1
      End
      Begin XtremeSuiteControls.ShortcutCaption sccFunBase 
         Height          =   360
         Left            =   0
         TabIndex        =   3
         Top             =   30
         Width           =   2505
         _Version        =   589884
         _ExtentX        =   4419
         _ExtentY        =   635
         _StockProps     =   6
         Caption         =   "基础设置"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
   End
   Begin VB.PictureBox picFunBack 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   4965
      Left            =   3510
      ScaleHeight     =   4965
      ScaleWidth      =   3555
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   3555
      Begin XtremeSuiteControls.ShortcutBar scbFunc 
         Height          =   4155
         Left            =   60
         TabIndex        =   1
         Top             =   210
         Width           =   3225
         _Version        =   589884
         _ExtentX        =   5689
         _ExtentY        =   7329
         _StockProps     =   64
      End
   End
   Begin MSComctlLib.ImageList imgPlan16 
      Left            =   9630
      Top             =   5880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   20
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicPlanMainFun.frx":04AA
            Key             =   "RootPlan"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicPlanMainFun.frx":07FC
            Key             =   "StopPlan"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicPlanMainFun.frx":705E
            Key             =   "TempletPlan"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicPlanMainFun.frx":75F8
            Key             =   "FixedPlan"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicPlanMainFun.frx":7B92
            Key             =   "InvalidFixedPlan"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicPlanMainFun.frx":812C
            Key             =   "InvalidPublishedFixedPlan"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicPlanMainFun.frx":86C6
            Key             =   "PublishedFixedPlan"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicPlanMainFun.frx":8C60
            Key             =   "InvalidMonthPlan"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicPlanMainFun.frx":91FA
            Key             =   "InvalidPublishedMonthPlan"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicPlanMainFun.frx":9794
            Key             =   "MonthPlan"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicPlanMainFun.frx":9D2E
            Key             =   "PublishedMonthPlan"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicPlanMainFun.frx":A2C8
            Key             =   "InvalidPublishedWeekPlan"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicPlanMainFun.frx":A862
            Key             =   "InvalidWeekPlan"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicPlanMainFun.frx":ADFC
            Key             =   "PublishedWeekPlan"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicPlanMainFun.frx":B396
            Key             =   "WeekPlan"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicPlanMainFun.frx":B930
            Key             =   "MonthTemplet"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicPlanMainFun.frx":BECA
            Key             =   "MonthTempletDay"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicPlanMainFun.frx":C464
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicPlanMainFun.frx":C9FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicPlanMainFun.frx":CF98
            Key             =   "WeekTemplet"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.ImageManager imgIcons 
      Left            =   2640
      Top             =   6000
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmClinicPlanMainFun.frx":D532
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   210
      Top             =   90
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmClinicPlanMainFun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmMain As Form
Private mcbsMain As Object            'CommandBar控件
Private mstrPrivs As String
Private mlngModule As Long

Private Enum ShortItemID
    ID_BaseItem = 10
    ID_PlanItem = 20
End Enum

Private mintCurYear As Integer '当前显示年份
Private mstrYear As String '可选年份，多个用"|"分隔
Private mblnNotClick As Boolean

Private mcllVisitTable As Collection 'Array(出诊ID,出诊表名,排班方式,模板类型)

Public Sub InitCommVariable(frmParent As Form, cbsThis As Object, _
    ByVal strPrivs As String, ByVal lngModule As Long)
    '初始化变量
    Set mfrmMain = frmParent
    Set mcbsMain = cbsThis
    mstrPrivs = strPrivs
    mlngModule = lngModule
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If Val(lblYearSelect.Tag) = Val(Control.Parameter) Then Exit Sub
    
    lblYearSelect.Caption = Val(Control.Parameter) & "年"
    lblYearSelect.Tag = Val(Control.Parameter)
    mintCurYear = Val(Control.Parameter)
    Call LoadVisitTable
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    Call mfrmMain.ActiveFormChange(Me)
End Sub

Private Sub Form_Load()
    Dim tpGroup As TaskPanelGroup
    Err = 0: On Error GoTo errHandler
    
    mintCurYear = Year(Now)
    lblYearSelect.Caption = mintCurYear & "年"
    With tplFunBase
        .Behaviour = xtpTaskPanelBehaviourList
        .HotTrackStyle = xtpTaskPanelHighlightItem
        .SelectItemOnFocus = True
        .Icons.AddIcons imgIcons.Icons
        .SetIconSize 32, 32
        .ItemLayout = xtpTaskItemLayoutImagesWithTextBelow
        .SetMargins 1, 0, 0, 1, 2
        
        Set tpGroup = .Groups.Add(10, "基础设置")
        tpGroup.Items.Add Pane_WorkTime, "上班时间管理", xtpTaskItemTypeLink, 12
        tpGroup.Items.Add Pane_Holiday, "节假日管理", xtpTaskItemTypeLink, 13
        tpGroup.Items.Add Pane_DoctorOffice, "门诊诊室管理", xtpTaskItemTypeLink, 14
        tpGroup.Items.Add Pane_SignalSource, "临床号源管理", xtpTaskItemTypeLink, 15
        
        tpGroup.CaptionVisible = False
        tpGroup.Expanded = True
    End With
    Call CreateShortcutBar
    Call LoadYear
    Call LoadVisitTable
    
    With scbFunc
        .Tag = ID_BaseItem
        mblnNotClick = True
        .Selected = .Item(1) '要切换一下，保证控件绑定到位\
        mblnNotClick = False
        .Selected = .Item(0)
    End With
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function LoadVisitTable() As Boolean
    '功能：加载出诊表列表
    '说明：由节点Key值格式区分出诊表类别
    '       停诊安排:K_StopPlan
    '       出诊表模板节点：K0_出诊ID
    '       固定安排的固定节点：K_FixedRoot
    '       固定出诊表节点：K1_出诊ID
    '       XX年出诊表节点：K_年份
    '       XX月出诊表节点：K2_年份_月份
    '       XX周出诊表节点：K3_年份_月份_周数
    Dim strSQL As String, strWhere As String, rsVisitTable As ADODB.Recordset
    Dim objYearNode As Node, objMonthNode As Node, objCurNode As Node
    Dim strKey As String, objNode As Node
    
    Err = 0: On Error GoTo errHandler
    Set mcllVisitTable = New Collection
    tvwPlanTemplet.Nodes.Clear
    tvwPlan.Nodes.Clear
    '如果存在"停诊申请"或"停诊审批"权限，在出诊安排中增加"停诊安排"节点
    If zlStr.IsHavePrivs(mstrPrivs, "停诊申请") Or zlStr.IsHavePrivs(mstrPrivs, "停诊审批") Then
        tvwPlan.Nodes.Add , , "K_StopPlan", "停诊安排", "StopPlan"
    End If
    tvwPlan.Nodes.Add , , "K_FixedRoot", "固定出诊", "RootPlan"
    
'    '没有所有科室表示临床排班
'    If zlStr.IsHavePrivs(mstrPrivs, "所有科室") = False Then
'        '根据挂号安排的号源去判断
'        strWhere = "And Exists" & vbNewLine & _
'                    "       (Select 1" & vbNewLine & _
'                    "         From 临床出诊安排 M, 临床出诊号源 N" & vbNewLine & _
'                    "         Where m.出诊id = a.Id And m.号源id + 0 = n.Id And Nvl(n.是否临床排班, 0) = 1" & vbNewLine & _
'                    "               And Exists(Select 1 From 部门人员 Where 部门id = n.科室id And 人员id = [2]))" & vbNewLine
'    End If
    strSQL = "Select a.ID, a.出诊表名, a.排班方式, a.年份, a.月份, a.周数," & vbNewLine & _
            "       Decode(a.发布时间, Null, 0, 1) As 是否发布,Nvl(模板类型,0) As 模板类型," & vbNewLine & _
            "       Nvl((Select 1 From 临床出诊安排 Where 出诊id = a.Id And 终止时间 >= Trunc(Sysdate) And Rownum < 2), 0) As 是否有效" & vbNewLine & _
            " From 临床出诊表 A" & vbNewLine & _
            " Where ((Nvl(a.排班方式, 0) = 3 And (nvl(a.应用范围,0)=2 or nvl(a.应用范围,0)=0 and a.发布人=[3]" & vbNewLine & _
            "        Or (Nvl(a.应用范围, 0) = 1 And a.科室id In (Select 部门id From 部门人员 Where 人员id = [2]))))" & vbNewLine & _
            "       Or (Nvl(a.排班方式, 0) In (1, 2) And a.年份 = [1]) Or Nvl(a.排班方式, 0) = 0)" & vbNewLine & _
            "       And Nvl(站点,'-')=Nvl([4],'-')" & vbNewLine & _
            strWhere & vbNewLine & _
            " Order By Decode(a.排班方式, 0, 0, 3, 3, 1), a.年份, a.月份, a.周数, a.ID"
    Set rsVisitTable = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mintCurYear, UserInfo.id, UserInfo.姓名, gstrNodeNo)
    If rsVisitTable Is Nothing Then GoTo ClearFixedRoot:
    If rsVisitTable.RecordCount = 0 Then GoTo ClearFixedRoot:
    
    '排班方式：0-固定排班;1-按月排班;2-按周排班;3-模板
    With rsVisitTable
        Do While Not .EOF
            'Array(出诊ID,出诊表名,排班方式,模板类型)
            mcllVisitTable.Add Array(Nvl(!id), Nvl(!出诊表名), Nvl(!排班方式), Val(Nvl(!模板类型))), "K0_" & Nvl(!id)
            If Nvl(!排班方式) = 3 Then  '模板,Key值格式：K0_出诊ID
                strKey = "K0_" & Nvl(!id)
                Set objNode = tvwPlanTemplet.Nodes.Add(, , strKey, Nvl(!出诊表名), Decode(Val(Nvl(!模板类型)), 1, "MonthTemplet", 2, "MonthTempletDay", "WeekTemplet"))
                objNode.Tag = Val(Nvl(!id))
            ElseIf Nvl(!排班方式) = 0 Then  '固定排班,Key值格式：K1_出诊ID
                strKey = "K1_" & Nvl(!id)
                If Not FindNodeByKey(tvwPlan.Nodes, "K_FixedRoot") Is Nothing Then
                    Set objNode = tvwPlan.Nodes("K_FixedRoot")
                    objNode.Text = Nvl(!出诊表名)
                    objNode.Key = strKey
                    objNode.Tag = Val(Nvl(!id))
                    objNode.Image = GetIconIndex(1, Val(Nvl(!是否发布)) = 1, Val(Nvl(!是否有效)) = 0)
                End If
            Else '月排班和周排班
                '1.年份节点,Key值格式：K_年份
                strKey = "K_" & Nvl(!年份)
                If FindNodeByKey(tvwPlan.Nodes, strKey) Is Nothing Then
                    Set objYearNode = tvwPlan.Nodes.Add(, , strKey, Nvl(!年份) & "年出诊安排", "RootPlan")
                Else
                    Set objYearNode = tvwPlan.Nodes(strKey)
                End If
                '2.月份节点,Key值格式：K2_年份_月份
                strKey = "K2_" & Nvl(!年份) & "_" & Nvl(!月份)
                If FindNodeByKey(tvwPlan.Nodes, strKey) Is Nothing Then
                    Set objMonthNode = tvwPlan.Nodes.Add(objYearNode, tvwChild, strKey, Nvl(!月份) & "月出诊表", "InvalidMonthPlan")
                    If Val(Nvl(!周数)) = 0 Then
                        objMonthNode.Tag = Val(Nvl(!id))
                        objMonthNode.Image = GetIconIndex(2, Val(Nvl(!是否发布)) = 1, Val(Nvl(!是否有效)) = 0)
                    End If
                Else
                    Set objMonthNode = tvwPlan.Nodes(strKey)
                    If Val(Nvl(!周数)) = 0 Then
                        objMonthNode.Tag = Val(Nvl(!id))
                        objMonthNode.Image = GetIconIndex(2, Val(Nvl(!是否发布)) = 1, Val(Nvl(!是否有效)) = 0)
                    End If
                End If
                '3.周数节点，Key值格式：K3_年份_月份_周数
                If Nvl(!排班方式) = 2 Then  '周排班
                    strKey = "K3_" & Nvl(!年份) & "_" & Nvl(!月份) & "_" & Nvl(!周数)
                    Set objNode = tvwPlan.Nodes.Add(objMonthNode, tvwChild, strKey, "第" & Nvl(!周数) & "周出诊表", "InvalidWeekPlan")
                    objNode.Tag = Val(Nvl(!id))
                    objNode.Image = GetIconIndex(3, Val(Nvl(!是否发布)) = 1, Val(Nvl(!是否有效)) = 0)
                End If
            End If
             .MoveNext
        Loop
    End With
    
    '展开节点
    For Each objCurNode In tvwPlan.Nodes
        objCurNode.Expanded = True
        
        '访问月出诊表节点，确定图标
        '只有所有子节点及自己都为无效时才显示为无效节点
        If InStr(objCurNode.Key, "_") > 0 Then
            If Split(objCurNode.Key, "_")(0) = "K2" Then
                If Val(objCurNode.Tag) <> 0 Then
                    '有月出诊表，则以月出诊表为准
                    '如果月出诊表无效，所有周出诊表肯定也无效
                Else
                    '只有周出诊表，无月出诊表
                    Set objNode = objCurNode.Child
                    Do While Not objNode Is Nothing
                        If objNode.Image = "WeekPlan" Or objCurNode.Image = "PublishedWeekPlan" Then
                            '一个周出诊表有效，则月出诊表节点有效
                            objCurNode.Image = "MonthPlan"
                        End If
                        Set objNode = objNode.Next
                    Loop
                End If
            End If
        End If
    Next
    
ClearFixedRoot:
    LoadVisitTable = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetIconIndex(ByVal bytMode As Byte, _
    Optional ByVal blnPublished As Boolean, Optional ByVal blnInvalid As Boolean) As String
    '获取出诊安排节点图标索引
    '入参：
    '   bytMode 1-固定出诊表,2-月出诊表,3-周出诊表
    Select Case bytMode
    Case 1
        If blnPublished Then
            If blnInvalid Then
                GetIconIndex = "InvalidPublishedFixedPlan"
            Else
                GetIconIndex = "PublishedFixedPlan"
            End If
        Else
            If blnInvalid Then
                GetIconIndex = "InvalidFixedPlan"
            Else
                GetIconIndex = "FixedPlan"
            End If
        End If
    Case 2
        If blnPublished Then
            If blnInvalid Then
                GetIconIndex = "InvalidPublishedMonthPlan"
            Else
                GetIconIndex = "PublishedMonthPlan"
            End If
        Else
            If blnInvalid Then
                GetIconIndex = "InvalidMonthPlan"
            Else
                GetIconIndex = "MonthPlan"
            End If
        End If
    Case 3
    If blnPublished Then
            If blnInvalid Then
                GetIconIndex = "InvalidPublishedWeekPlan"
            Else
                GetIconIndex = "PublishedWeekPlan"
            End If
        Else
            If blnInvalid Then
                GetIconIndex = "InvalidWeekPlan"
            Else
                GetIconIndex = "WeekPlan"
            End If
        End If
    Case Else
        GetIconIndex = "FixedPlan"
    End Select
End Function

Private Sub CreateShortcutBar()
    Err = 0: On Error GoTo errHandler
    With scbFunc
        .Icons = imgIcons.Icons
        
        '图像索引号与ID相同
        .AddItem ID_PlanItem, "出诊安排", picPlanBack.Hwnd
        .AddItem ID_BaseItem, "基础设置", picBaseSetBack.Hwnd
        .ExpandedLinesCount = .ItemCount '默认展开
    End With
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    picFunBack.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub frmMoveSplitY_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Err = 0: On Error Resume Next
    If Button <> vbLeftButton Then Exit Sub
    If tvwPlanTemplet.Height + Y < 1200 Or tvwPlan.Height - Y < 1500 Then Exit Sub
    
    tvwPlanTemplet.Height = tvwPlanTemplet.Height + Y
    frmMoveSplitY.Top = frmMoveSplitY.Top + Y
    sccPlan.Top = sccPlan.Top + Y
    lblYearSelect.Top = sccPlan.Top + (sccPlan.Height - lblYearSelect.Height) / 2
    imgYearSelect.Top = sccPlan.Top + (sccPlan.Height - imgYearSelect.Height) / 2
    tvwPlan.Top = tvwPlan.Top + Y
    tvwPlan.Height = tvwPlan.Height - Y
End Sub

Private Sub imgYearSelect_Click()
    Call lblYearSelect_Click
End Sub

Private Sub lblYearSelect_Click()
    Call ShowPopuYear
End Sub

Public Function RefreshVisitTable(Optional ByVal strKey As String) As Boolean
    '刷新出诊表
    Dim objNode As Node
    Dim strDeletedPreviouNodeKey As String
    
    Err = 0: On Error GoTo errHandler
    If scbFunc.Selected.id <> ID_PlanItem Then Exit Function
    
    If strKey = "" Then
        If Me.ActiveControl Is tvwPlanTemplet Then
            If Not tvwPlanTemplet.SelectedItem Is Nothing Then Set objNode = tvwPlanTemplet.SelectedItem
        Else
            If Not tvwPlan.SelectedItem Is Nothing Then Set objNode = tvwPlan.SelectedItem
        End If
        '删除节点时，确定选中节点的Key值
        If Not objNode Is Nothing Then
            strKey = objNode.Key
            If Not objNode.Previous Is Nothing Then
                strDeletedPreviouNodeKey = objNode.Previous.Key
            ElseIf Not objNode.Next Is Nothing Then
                strDeletedPreviouNodeKey = objNode.Next.Key
            ElseIf Not objNode.Parent Is Nothing Then
                strDeletedPreviouNodeKey = objNode.Parent.Key
            End If
        End If
    Else
        If Left(strKey, 2) = "K0" Then
            If tvwPlanTemplet.Visible And tvwPlanTemplet.Enabled Then tvwPlanTemplet.SetFocus
        Else
            If tvwPlan.Visible And tvwPlan.Enabled Then tvwPlan.SetFocus
        End If
    End If
    
    Call LoadYear
    Call LoadVisitTable '重新加载出诊表
    
    '模板
    If Me.ActiveControl Is tvwPlanTemplet Then
        '先定位到被选中项
        Set objNode = FindNodeByKey(tvwPlanTemplet.Nodes, strKey)
        If Not objNode Is Nothing Then
            tvwPlanTemplet.Tag = ""
            objNode.Selected = True
            tvwPlanTemplet_NodeClick objNode
            tvwPlanTemplet.SetFocus
            RefreshVisitTable = True: Exit Function
        Else
            '定位到上一个
            Set objNode = FindNodeByKey(tvwPlanTemplet.Nodes, strDeletedPreviouNodeKey)
            If Not objNode Is Nothing Then
                tvwPlanTemplet.Tag = ""
                objNode.Selected = True
                tvwPlanTemplet_NodeClick objNode
                tvwPlanTemplet.SetFocus
                RefreshVisitTable = True: Exit Function
            End If
        End If
    Else
        '出诊安排
        Set objNode = FindNodeByKey(tvwPlan.Nodes, strKey)
        If Not objNode Is Nothing Then
            tvwPlan.Tag = ""
            objNode.Selected = True
            tvwPlan_NodeClick objNode
            tvwPlan.SetFocus
            RefreshVisitTable = True: Exit Function
        Else
            Set objNode = FindNodeByKey(tvwPlan.Nodes, strDeletedPreviouNodeKey)
            If Not objNode Is Nothing Then
                tvwPlan.Tag = ""
                objNode.Selected = True
                tvwPlan_NodeClick objNode
                tvwPlan.SetFocus
                RefreshVisitTable = True: Exit Function
            End If
        End If
    End If
    
    Call tvwPlan_GotFocus
    RefreshVisitTable = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub picBaseSetBack_Resize()
    Err = 0: On Error Resume Next
    sccFunBase.Move 0, -10, picBaseSetBack.ScaleWidth
    With tplFunBase
        .Left = 0
        .Top = sccFunBase.Top + sccFunBase.Height
        .Width = picBaseSetBack.ScaleWidth - .Left
        .Height = picBaseSetBack.ScaleHeight - .Top
    End With
End Sub

Private Sub picPlanBack_Resize()
    Err = 0: On Error Resume Next
    sccPlanTemplet.Move 0, -10, picPlanBack.ScaleWidth
    tvwPlanTemplet.Move 0, sccPlanTemplet.Top + sccPlanTemplet.Height, picPlanBack.ScaleWidth
    frmMoveSplitY.Move -25, tvwPlanTemplet.Top + tvwPlanTemplet.Height, picPlanBack.ScaleWidth + 100
    
    sccPlan.Move 0, frmMoveSplitY.Top + frmMoveSplitY.Height, picPlanBack.ScaleWidth
    imgYearSelect.Top = sccPlan.Top + (sccPlan.Height - imgYearSelect.Height) / 2
    imgYearSelect.Left = sccPlan.Width - imgYearSelect.Width - 10
    lblYearSelect.Top = sccPlan.Top + (sccPlan.Height - lblYearSelect.Height) / 2
    lblYearSelect.Left = imgYearSelect.Left - lblYearSelect.Width - 10
    tvwPlan.Move 0, sccPlan.Top + sccPlan.Height, picPlanBack.ScaleWidth, picPlanBack.ScaleHeight - (sccPlan.Top + sccPlan.Height)
End Sub

Private Sub picFunBack_Resize()
    Err = 0: On Error Resume Next
    scbFunc.Move 0, 0, picFunBack.ScaleWidth, picFunBack.ScaleHeight
End Sub

Private Sub scbFunc_ExpandButtonDown(CancelMenu As Boolean)
    CancelMenu = True
End Sub

Private Sub scbFunc_SelectedChanged(ByVal Item As XtremeSuiteControls.IShortcutBarItem)
    Dim tpGroup As TaskPanelGroup
    Dim tpItem As TaskPanelGroupItem
    Dim blnFind As Boolean, tpItemWork As TaskPanelGroupItem
    
    Err = 0: On Error GoTo errHandler
    If mblnNotClick Then Exit Sub
    If Val(scbFunc.Tag) = Item.id Then Exit Sub
    
    tvwPlanTemplet.Tag = ""
    tvwPlan.Tag = ""
    
    '设置默认选中节点
    picBaseSetBack.Visible = False
    picPlanBack.Visible = False
    If Item.id = ID_BaseItem Then
        scbFunc.Tag = ID_BaseItem
        If tplFunBase.Tag = "" Then tplFunBase.Tag = "临床号源管理" '缺省选择“临床号源管理”
        For Each tpGroup In tplFunBase.Groups
            For Each tpItem In tpGroup.Items
                If tpItem.Caption = tplFunBase.Tag Then Set tpItemWork = tpItem
                If tplFunBase.Tag = tpItem.Caption Then
                    tpItem.Selected = True: blnFind = True
                    tplFunBase.Tag = "": tplFunBase_ItemClick tpItem
                Else
                    tpItem.Selected = False
                End If
            Next
        Next
        If blnFind = False Then
            '缺省选择上班时段
            tpItemWork.Selected = True
            tplFunBase.Tag = "": tplFunBase_ItemClick tpItemWork
        End If
        picBaseSetBack.Visible = True
    Else
        scbFunc.Tag = ID_PlanItem
        Call tvwPlan_GotFocus
        
        picPlanBack.Visible = True
    End If
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub tplFunBase_ItemClick(ByVal Item As XtremeSuiteControls.ITaskPanelGroupItem)
    Err = 0: On Error GoTo errHandler
    If tplFunBase.Tag = Item.Caption Then Exit Sub
    tplFunBase.Tag = Item.Caption
    
    Select Case Item.Caption
    Case "上班时间管理"
        Call mfrmMain.SelectedChange(Pane_WorkTime)
    Case "节假日管理"
        Call mfrmMain.SelectedChange(Pane_Holiday)
    Case "门诊诊室管理"
        Call mfrmMain.SelectedChange(Pane_DoctorOffice)
    Case "临床号源管理"
        Call mfrmMain.SelectedChange(Pane_SignalSource)
    End Select
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub tvwPlan_GotFocus()
    Err = 0: On Error GoTo errHandler
    
    tvwPlanTemplet.Tag = ""
    tvwPlanTemplet.HideSelection = True
    tvwPlan.HideSelection = False
    
    If tvwPlan.Nodes.Count > 0 Then
        If tvwPlan.SelectedItem Is Nothing Then
            tvwPlan.Tag = ""
            If tvwPlan.Nodes.Count > 1 Then
                tvwPlan.Nodes(2).Selected = True
            Else
                tvwPlan.Nodes(1).Selected = True
            End If
        End If
        tvwPlan_NodeClick tvwPlan.SelectedItem
    Else
        Call mfrmMain.SelectedChange(Pane_FixedPlan)
    End If
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub tvwPlan_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '弹出右键菜单
    Dim cbCommandBar As CommandBar
    
    Err = 0: On Error GoTo errHandler
    If Not (Button = vbRightButton) Then Exit Sub
    If Not (tvwPlan.Visible And tvwPlan.Enabled) Then Exit Sub
    tvwPlan.SetFocus: Call mfrmMain.ActiveFormChange(Me)
    Call tvwPlan_GotFocus
    
    Set cbCommandBar = mfrmMain.GetPopupCommandBarSub()
    If cbCommandBar Is Nothing Then Exit Sub
    If cbCommandBar.Controls.Count = 0 Then Exit Sub
    
    cbCommandBar.ShowPopup
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub tvwPlan_NodeClick(ByVal Node As MSComctlLib.Node)
    Err = 0: On Error GoTo errHandler
    If tvwPlan.Tag = Node.Key Then Exit Sub

    tvwPlanTemplet.Tag = ""
    tvwPlan.Tag = Node.Key
    tvwPlan.HideSelection = False

    Select Case Split(Node.Key, "_")(0)
    Case "K1" '固定安排
        Call mfrmMain.SelectedChange(Pane_FixedPlan, Val(Node.Tag))
    Case "K2" '月排班
        Call mfrmMain.SelectedChange(Pane_MonthPlan, Val(Node.Tag), _
            Val(Split(Node.Key, "_")(1)), Val(Split(Node.Key, "_")(2)), _
            Val(Split(Node.Key, "_")(1)) & "年" & Val(Split(Node.Key, "_")(2)) & "月")
    Case "K3" '周排班
        Call mfrmMain.SelectedChange(Pane_WeekPlan, Val(Node.Tag))
    Case Else
        If Node.Key = "K_StopPlan" Then '停诊管理
            Call mfrmMain.SelectedChange(Pane_StopPlan)
        ElseIf Node.Parent Is Nothing Then
            If Node.Key = "K_FixedRoot" Then
                Call mfrmMain.SelectedChange(Pane_FixedPlan, Val(Node.Tag))
            Else
                Call mfrmMain.SelectedChange(Pane_WeekPlan, Val(Node.Tag), 0, 0, "出诊表")
            End If
        End If
    End Select
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub tvwPlanTemplet_GotFocus()
    Err = 0: On Error GoTo errHandler
    
    tvwPlan.Tag = ""
    tvwPlanTemplet.HideSelection = False
    tvwPlan.HideSelection = True
    
    If tvwPlanTemplet.Nodes.Count > 0 Then
        If tvwPlanTemplet.SelectedItem Is Nothing Then
            tvwPlanTemplet.Tag = ""
            tvwPlanTemplet.Nodes(1).Selected = True
        End If
        tvwPlanTemplet_NodeClick tvwPlanTemplet.SelectedItem
    Else
        Call mfrmMain.SelectedChange(Pane_PlanTemplet)
    End If
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub tvwPlanTemplet_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '弹出右键菜单
    Dim cbCommandBar As CommandBar
    
    Err = 0: On Error GoTo errHandler
    If Not (Button = vbRightButton) Then Exit Sub
    If Not (tvwPlanTemplet.Visible And tvwPlanTemplet.Enabled) Then Exit Sub
    tvwPlanTemplet.SetFocus: Call mfrmMain.ActiveFormChange(Me)
    Call tvwPlanTemplet_GotFocus
    
    Set cbCommandBar = mfrmMain.GetPopupCommandBarSub()
    If cbCommandBar Is Nothing Then Exit Sub
    If cbCommandBar.Controls.Count = 0 Then Exit Sub
    
    cbCommandBar.ShowPopup
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub tvwPlanTemplet_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim bytTempletType As Byte
    
    Err = 0: On Error GoTo errHandler
    If tvwPlanTemplet.Tag = Node.Key Then Exit Sub
    
    tvwPlan.Tag = ""
    tvwPlanTemplet.Tag = Node.Key
    tvwPlanTemplet.HideSelection = False
    
    'Array(出诊ID,出诊表名,排班方式,模板类型)
    If CollExitsValue(mcllVisitTable, Node.Key) Then
        bytTempletType = Val(mcllVisitTable(Node.Key)(3))
    End If
    If bytTempletType = 2 Then
        Call mfrmMain.SelectedChange(Pane_MonthTemplet, Val(Node.Tag))
    Else
        Call mfrmMain.SelectedChange(Pane_PlanTemplet, Val(Node.Tag), 0, 0, "", bytTempletType)
    End If
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub LoadYear()
    '加载可选年份
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strYear As String
    Dim blnFind As Boolean
    
    Err = 0: On Error GoTo errHandler
    mstrYear = ""
    strSQL = "Select Distinct 年份 From 临床出诊表 Where 排班方式 In (1, 2) And Nvl(站点,'-') = Nvl([1],'-') Order By 年份"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, gstrNodeNo)
    If Not rsTemp.EOF Then
        Do While Not rsTemp.EOF
            mstrYear = mstrYear & "|" & Nvl(rsTemp!年份)
            If Val(Nvl(rsTemp!年份)) = mintCurYear Then blnFind = True
            rsTemp.MoveNext
        Loop
    End If
    If mstrYear = "" Then
        mstrYear = mintCurYear
    Else
        mstrYear = Mid(mstrYear, 2)
    End If
    If blnFind = False Then
        mintCurYear = Split(mstrYear, "|")(UBound(Split(mstrYear, "|")))
        lblYearSelect.Caption = mintCurYear & "年"
    End If
    Exit Sub
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function CreatePopuMenu(ByVal strYear As String) As CommandBar
    '功能:创建临时菜单
    Dim objCommandBar As CommandBar
    Dim objControl As CommandBarControl
    Dim i As Integer, varYear As Variant
    
    If strYear = "" Then Exit Function
    
    cbsThis.DeleteAll
    Set objCommandBar = cbsThis.Add("PopupYear", xtpBarPopup)
    With objCommandBar.Controls
        varYear = Split(strYear, "|")
        For i = 0 To UBound(varYear)
            Set objControl = .Add(xtpControlButton, 1000 + i, Val(varYear(i)) & "年")
            objControl.Parameter = Val(varYear(i))
            If Val(varYear(i)) = mintCurYear Then
                objControl.Checked = True
            End If
        Next
    End With
    Set CreatePopuMenu = objCommandBar
    Set objCommandBar = Nothing
End Function

Private Sub ShowPopuYear()
    Dim objCommandBar As CommandBar
    Dim vRect As RECT
    
    vRect = zlControl.GetControlRect(picPlanBack.Hwnd)
    vRect.Left = vRect.Left + lblYearSelect.Left - 2
    vRect.Top = vRect.Top + lblYearSelect.Top + 2
    Set objCommandBar = CreatePopuMenu(mstrYear)
    If objCommandBar Is Nothing Then Exit Sub
    
    Call objCommandBar.ShowPopup(, vRect.Left, vRect.Top + lblYearSelect.Height)
    Set objCommandBar = Nothing
End Sub


