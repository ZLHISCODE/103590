VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRequestNavigation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "卫材申领管理自动生成向导"
   ClientHeight    =   5070
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   7890
   Icon            =   "frmRequestNavigation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   7890
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   45
      TabIndex        =   29
      Top             =   4590
      Width           =   1100
   End
   Begin VB.PictureBox PicSetup 
      Height          =   4485
      Left            =   0
      ScaleHeight     =   4425
      ScaleWidth      =   1425
      TabIndex        =   3
      Top             =   -15
      Width           =   1485
      Begin VB.Image imgSetup 
         Height          =   4335
         Left            =   60
         Picture         =   "frmRequestNavigation.frx":1582
         Stretch         =   -1  'True
         Top             =   60
         Width           =   1320
      End
   End
   Begin VB.CommandButton cmdPrevious 
      Caption         =   "上一步(&B)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   3675
      TabIndex        =   1
      Top             =   4605
      Width           =   1230
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6540
      TabIndex        =   2
      Top             =   4605
      Width           =   1230
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "下一步(&N)"
      Default         =   -1  'True
      Height          =   350
      Left            =   5040
      TabIndex        =   0
      Top             =   4605
      Width           =   1230
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRequestNavigation.frx":6B68
            Key             =   "Item"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraStep 
      Height          =   4605
      Index           =   0
      Left            =   1470
      TabIndex        =   4
      Top             =   -120
      Width           =   6435
      Begin VB.Frame FraNote 
         Height          =   30
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   660
         Width           =   6255
      End
      Begin VB.ComboBox cbo库房 
         Height          =   300
         Left            =   2130
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1140
         Width           =   2235
      End
      Begin VB.Frame Frame1 
         Caption         =   "申领方式"
         Height          =   2865
         Index           =   0
         Left            =   150
         TabIndex        =   10
         Top             =   1530
         Width           =   6105
         Begin VB.OptionButton optMode 
            Caption         =   "根据指定时间范围内的申领单汇总"
            Height          =   195
            Index           =   4
            Left            =   330
            TabIndex        =   30
            Top             =   2364
            Width           =   3516
         End
         Begin VB.OptionButton optMode 
            Caption         =   "根据指定时间范围内的申领单"
            Height          =   195
            Index           =   3
            Left            =   330
            TabIndex        =   14
            Top             =   1827
            Width           =   2685
         End
         Begin VB.OptionButton optMode 
            Caption         =   "按材料的储备下限"
            Height          =   195
            Index           =   2
            Left            =   330
            TabIndex        =   13
            Top             =   1293
            Width           =   2685
         End
         Begin VB.OptionButton optMode 
            Caption         =   "按材料的储备上限"
            Height          =   195
            Index           =   1
            Left            =   330
            TabIndex        =   12
            Top             =   759
            Width           =   2685
         End
         Begin VB.OptionButton optMode 
            Caption         =   "根据指定时间范围内材料的消耗量"
            Height          =   180
            Index           =   0
            Left            =   330
            TabIndex        =   11
            Top             =   240
            Value           =   -1  'True
            Width           =   3045
         End
         Begin VB.Label lblTip 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "根据指定时间范围内的申领单汇总，产生本次申领单"
            ForeColor       =   &H00004000&
            Height          =   180
            Index           =   4
            Left            =   396
            TabIndex        =   31
            Top             =   2556
            Width           =   4140
         End
         Begin VB.Label lblTip 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "根据指定时间范围内的申领单的未发数量，产生本次申领单"
            ForeColor       =   &H00004000&
            Height          =   180
            Index           =   3
            Left            =   396
            TabIndex        =   18
            Top             =   2027
            Width           =   4680
         End
         Begin VB.Label lblTip 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "使当前库房的材料储备量始终保持在下限标准，产生本次申领单"
            ForeColor       =   &H00004000&
            Height          =   180
            Index           =   2
            Left            =   396
            TabIndex        =   17
            Top             =   1498
            Width           =   5040
         End
         Begin VB.Label lblTip 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "使当前库房的材料储备量始终保持在上限标准，产生本次申领单"
            ForeColor       =   &H00004000&
            Height          =   180
            Index           =   1
            Left            =   396
            TabIndex        =   16
            Top             =   969
            Width           =   5040
         End
         Begin VB.Label lblTip 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "根据您指定的时间范围，以材料的消耗量为依据，产生本次的申领单"
            ForeColor       =   &H00004000&
            Height          =   180
            Index           =   0
            Left            =   396
            TabIndex        =   15
            Top             =   440
            Width           =   5400
         End
      End
      Begin VB.Label lblCaption 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "第一步：决定产生申领单的方式"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   285
         Index           =   0
         Left            =   150
         TabIndex        =   9
         Top             =   240
         Width           =   4200
      End
      Begin VB.Label lblNote 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "你准备向哪个库房发生申领请求"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   180
         Left            =   1650
         TabIndex        =   7
         Top             =   900
         Width           =   2730
      End
      Begin VB.Label lbl库房 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "库房"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   1680
         TabIndex        =   5
         Top             =   1200
         Width           =   360
      End
   End
   Begin VB.Frame fraStep 
      Height          =   4605
      Index           =   1
      Left            =   1470
      TabIndex        =   19
      Top             =   -120
      Width           =   6435
      Begin MSComCtl2.DTPicker dtp开始时间 
         Height          =   285
         Left            =   4170
         TabIndex        =   23
         Top             =   1350
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   52101123
         CurrentDate     =   38096
      End
      Begin VB.Frame FraNote 
         Height          =   30
         Index           =   1
         Left            =   120
         TabIndex        =   20
         Top             =   660
         Width           =   6255
      End
      Begin MSComCtl2.DTPicker dtp结束时间 
         Height          =   285
         Left            =   4170
         TabIndex        =   25
         Top             =   1980
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   52101123
         CurrentDate     =   38096
      End
      Begin MSComctlLib.TreeView tvw分类 
         Height          =   3465
         Left            =   90
         TabIndex        =   28
         Top             =   1050
         Width           =   3840
         _ExtentX        =   6773
         _ExtentY        =   6112
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   494
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         ImageList       =   "img16"
         Appearance      =   1
      End
      Begin VB.Label lbl其它条件设置 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "其它条件设置"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4170
         TabIndex        =   27
         Top             =   810
         Width           =   1170
      End
      Begin VB.Label lbl分类选择 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "分类选择"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   180
         TabIndex        =   26
         Top             =   810
         Width           =   780
      End
      Begin VB.Label lbl结束时间 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "结束时间(&E)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   180
         Left            =   4200
         TabIndex        =   24
         Top             =   1740
         Width           =   1095
      End
      Begin VB.Label lbl开始时间 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "开始时间(&S)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   180
         Left            =   4200
         TabIndex        =   22
         Top             =   1110
         Width           =   1095
      End
      Begin VB.Label lblCaption 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "第二步：按指定分类进行条件搜索"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   285
         Index           =   1
         Left            =   150
         TabIndex        =   21
         Top             =   240
         Width           =   4500
      End
   End
End
Attribute VB_Name = "frmRequestNavigation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum 模式
    根据消耗量
    根据上限
    根据下限
    根据申领单未发数
    根据申领单汇总
End Enum
Private mblnOk As Boolean
Private mlngStockID As Long                 '申领库房ID
Private mbln明确批次 As Boolean             '申领时是否明确批次
Private mintCheck As Integer                '库存检查参数
Private mblnFirst  As Boolean
Private mintUnit As Integer
Private Const mlngModule = 1722

'----------------------------------------------------------------------------------------------------------
'刘兴宏:增加小数位数的格式串
'修改:2007/03/06
Private mFMT As g_FmtString
'----------------------------------------------------------------------------------------------------------

Private Const mconIntCol药名  As Integer = 2
Private Const mconIntCol规格   As Integer = 3
Private Const mconIntCol序号    As Integer = 4
Private Const mconIntCol分批核算  As Integer = 5
Private Const mconIntCol最大效期  As Integer = 6
Private Const mconIntCol可用数量  As Integer = 7
Private Const mconIntCol指导差价率 As Integer = 8
Private Const mconIntCol实际金额 As Integer = 9
Private Const mconIntCol实际差价 As Integer = 10
Private Const mconIntCol比例系数 As Integer = 11
Private Const mconIntCol批次 As Integer = 12
Private Const mconIntCol产地 As Integer = 13
Private Const mconIntCol批准文号 As Integer = 14
Private Const mconIntCol单位 As Integer = 15
Private Const mconIntCol批号 As Integer = 16
Private Const mconIntCol效期 As Integer = 17
Private Const mconIntCol灭菌失效期 As Integer = 18
Private Const mconIntCol填写数量 As Integer = 21
Private Const mconIntCol实际数量 As Integer = 22
Private Const mconIntCol采购价 As Integer = 24
Private Const mconIntCol采购金额 As Integer = 25
Private Const mconIntCol售价 As Integer = 26
Private Const mconIntCol售价金额 As Integer = 27
Private Const mconintCol差价 As Integer = 28
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int(glngSys / 100))
End Sub

Private Sub cmdNext_Click()
    
    If fraStep(0).Visible Then
        fraStep(1).Visible = True
        fraStep(0).Visible = False
        fraStep(1).ZOrder
        cmdPrevious.Enabled = True
        cmdNext.Caption = "完成(&F)"
        
        Call ResizeStuff
    Else
        '确定相应分类:
        Dim i As Long
        Dim str分类ID As String
        str分类ID = ""
        For i = 1 To tvw分类.Nodes.count
            If tvw分类.Nodes(i).Key <> "Root" And _
                tvw分类.Nodes(i).Checked Then
                str分类ID = str分类ID & "," & Mid(tvw分类.Nodes(i).Key, 2)
            End If
        Next
        
        If str分类ID <> "" Then
            str分类ID = Mid(str分类ID, 2)
        End If
    
        If Not CheckData(str分类ID) Then Exit Sub
        
        mblnOk = True
        Unload Me
    End If
End Sub

Private Sub cmdPrevious_Click()
    If fraStep(1).Visible Then
        fraStep(1).Visible = False
        fraStep(0).Visible = True
        fraStep(0).ZOrder
        cmdPrevious.Enabled = False
        cmdNext.Caption = "下一步(&N)"
    End If
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    '----缺省选中所有剂型----
    If Not mblnFirst Then Exit Sub
    fraStep(0).ZOrder
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim rsTemp As New ADODB.Recordset
    Dim strReg As String
    
    On Error GoTo ErrHandle
    strReg = Val(zlDatabase.GetPara("卫材单位", glngSys, mlngModule, "0"))
    mintUnit = Val(strReg)
  
    '刘兴宏:增加小数格式化串
    With mFMT
        .FM_成本价 = GetFmtString(mintUnit, g_成本价)
        .FM_金额 = GetFmtString(mintUnit, g_金额)
        .FM_零售价 = GetFmtString(mintUnit, g_售价)
        .FM_数量 = GetFmtString(mintUnit, g_数量)
    End With
    
    
    mblnFirst = True
    
    '----提取卫材库房----
    Set rsTemp = ReturnSQL(mlngStockID, "提取可以申领的库房", False, , 1722)
    
    If rsTemp.EOF Then
        MsgBox "没有任何库房允许申领，请在[基础参数设置]的卫材流向中设置！", vbInformation, gstrSysName
        Unload Me
        Exit Sub
    End If
    With cbo库房
        .Clear
        Do While Not rsTemp.EOF
            .AddItem rsTemp!名称
            .ItemData(.NewIndex) = rsTemp!Id
            rsTemp.MoveNext
        Loop
        .ListIndex = 0
    End With
    
    '----判断是否需要明确批次----
    mbln明确批次 = IS批次申领
    
    
   gstrSQL = "" & _
        "   Select Level as 层,ID,上级ID,名称 From 诊疗分类目录 where 类型=7" & _
        "   Start With 上级ID is NULL Connect by Prior ID=上级ID" & _
        "   Order by 层"
    
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
    If rsTemp.RecordCount = 0 Then
        MsgBox "卫材分类不完整！", vbInformation, gstrSysName
        Exit Sub
    End If

    Dim objNode As Node
    Set objNode = tvw分类.Nodes.Add(, , "Root", "所有卫材分类", "Item")
    
    Do While Not rsTemp.EOF
        If rsTemp!层 = 1 Then
            Set objNode = tvw分类.Nodes.Add("Root", 4, "_" & rsTemp!Id, rsTemp!名称, "Item")
        Else
            Set objNode = tvw分类.Nodes.Add("_" & rsTemp!上级ID, 4, "_" & rsTemp!Id, rsTemp!名称, "Item")
        End If
        rsTemp.MoveNext
    Loop
    tvw分类.Nodes("Root").Selected = True
    tvw分类.Nodes("Root").Expanded = True
    '----设置缺省的时间范围（一个月）----
    Me.dtp开始时间.Value = Format(DateAdd("d", -7, sys.Currentdate()), "yyyy-MM-dd") & " 00:00:00"
    Me.dtp结束时间.Value = Format(sys.Currentdate(), "yyyy-MM-dd HH:mm:ss")
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub ResizeStuff()
    Dim blnEnable As Boolean
    '判断是否允许用户输入其它条件
    blnEnable = (optMode(根据申领单未发数) Or optMode(根据消耗量) Or optMode(根据申领单汇总))
    lbl其它条件设置.Visible = blnEnable
    lbl开始时间.Visible = blnEnable
    lbl结束时间.Visible = blnEnable
    dtp开始时间.Visible = blnEnable
    dtp结束时间.Visible = blnEnable
    
    If blnEnable Then
        tvw分类.Width = lbl开始时间.Left - 200 - tvw分类.Left
    Else
        tvw分类.Width = fraStep(1).Width - 200 - tvw分类.Left
    End If
End Sub

Public Function ShowNavigation(ByVal frmParent As Object, ByVal lngStockID As Long) As Boolean
    On Error Resume Next
    mlngStockID = lngStockID
    mblnOk = False
    Me.Show 1, frmParent
    ShowNavigation = mblnOk
End Function

Private Function CheckData(Optional str分类id_IN As String = "") As Boolean
    Dim lngTargetID As Long             '目标库房的ID
    Dim rsCheck As New ADODB.Recordset
    Dim str分类IN As String
    
    '检查是否存在符合条件的记录（始终只按总数量进行比较，具体分的时候再按是否明确批次来分配各批次）
    On Error GoTo ErrHand
    CheckData = False
    
    gstrSQL = ""
    lngTargetID = cbo库房.ItemData(cbo库房.ListIndex)
    
    If optMode(根据消耗量) Then
        '如果明确批次，则药品库存中没有记录的卫材数据，不提取出来
     gstrSQL = "" & _
                 " Select sum(Nvl(A.实际数量,0)) 申领数量,max(Nvl(B.可用数量,0)) 可用数量,max(Nvl(B.实际数量,0)) 实际数量,max(Nvl(B.实际金额,0)) 实际金额,max(Nvl(B.实际差价,0)) 实际差价, " & _
                 "        D.材料ID,F.编码,F.名称,F.是否变价,D.库房分批,D.在用分批,P.现价 售价,F.规格,F.产地,D.最大效期,D.指导差价率, " & _
                 "        D.包装单位,D.换算系数,F.计算单位 售价单位 " & _
                 " From 药品收发记录 A,材料特性 D,收费项目目录 F,诊疗执行科室 Z,诊疗项目目录 N, " & _
                 "      (Select 收费细目ID,现价 From 收费价目 Where SysDate Between 执行日期 And Nvl(终止日期,Sysdate)" & _
                 GetPriceClassString("") & ") P, " & _
                 "      (Select 药品id 材料ID,sum(可用数量) as 可用数量,sum(实际数量) 实际数量,sum(实际金额) 实际金额,sum(实际差价) 实际差价" & _
                 "       From 药品库存 Where 库房ID=[4] And 性质=1 Group by 药品id) B " & IIf(str分类id_IN = "", "", ",Table(Cast(f_Num2List([5]) As zlTools.t_NumList)) Q") & _
                 " Where A.单据 IN (20,21,24,25,26) And A.入出系数=-1 And A.审核日期 Between [2]" & _
                 " And [3]" & _
                 " And A.药品ID=D.材料ID and D.诊疗ID=N.id  And Nvl(A.药品id,0)=B.材料ID(+) And A.库房ID=[1] " & _
                 " And A.药品ID=D.材料ID  And D.材料ID=P.收费细目ID" & IIf(str分类id_IN = "", "", " And N.分类id+0=Q.Column_Value") & _
                 " And D.材料ID=F.ID And D.诊疗ID=Z.诊疗项目ID And Z.执行科室ID+0=[1]" & _
                 " Having Sum(Nvl(A.实际数量,0))>0 " & _
                 " Group By D.材料ID,F.编码,F.名称,F.是否变价,D.库房分批,D.在用分批,P.现价,F.规格,F.产地,D.最大效期,D.指导差价率, " & _
                 "       D.包装单位,D.换算系数,F.计算单位 "
                 
    ElseIf optMode(根据上限) Then
       gstrSQL = "Select Nvl(A.上限,0)-Sum(Nvl(B.可用数量,0)) 申领数量,Sum(Nvl(K.可用数量,0)) 可用数量,Sum(Nvl(K.实际数量,0)) 实际数量,Sum(Nvl(K.实际金额,0)) 实际金额,Sum(Nvl(K.实际差价,0)) 实际差价,  " & _
                "         D.材料ID,F.编码,F.名称,F.是否变价,D.库房分批,D.在用分批,P.现价 售价,F.规格,F.产地,D.最大效期,D.指导差价率,  " & _
                "         D.包装单位,D.换算系数,F.计算单位 售价单位  " & _
                "  From (Select 库房id, 材料id, 上限, 下限, 盘点属性, 库房货位 From 材料储备限额 Where 库房ID=[1] And Nvl(上限,0)>0) A, " & _
                "       材料特性 D,收费项目目录 F,诊疗执行科室 Z,诊疗项目目录 N,  " & _
                "       (Select 收费细目ID,现价 From 收费价目 Where SysDate Between 执行日期 And Nvl(终止日期,Sysdate)" & _
                GetPriceClassString("") & ") P,  " & _
                "       (Select 药品id 材料ID,Sum(可用数量) as 可用数量,sum(实际数量) 实际数量, sum(实际金额) 实际金额,sum(实际差价) 实际差价 " & _
                "        From 药品库存 Where 库房ID=[1] And 性质=1 Group by 药品id) B,  " & _
                 "      (Select 药品id 材料ID,Sum(可用数量) as 可用数量,sum(实际数量) 实际数量, sum(实际金额) 实际金额,sum(实际差价) 实际差价 " & _
                 "      From 药品库存 Where 库房ID=[4] And 性质=1 Group by 药品id ) K " & IIf(str分类id_IN = "", "", ",Table(Cast(f_Num2List([5]) As zlTools.t_NumList)) Q") & _
                "  Where A.材料ID=D.材料ID  and A.材料id=b.材料id(+) And A.材料id=K.材料ID(+) And D.材料ID=P.收费细目ID" & IIf(str分类id_IN = "", "", " And N.分类id+0=Q.Column_Value") & _
                "  And D.材料ID=F.ID and D.诊疗ID=N.id  And D.诊疗ID=Z.诊疗项目ID And Z.执行科室ID+0=[1]" & _
                "  Having Nvl(A.上限,0)-Sum(Nvl(B.可用数量,0))>0 " & _
                "  Group By Nvl(A.上限,0),D.材料ID,F.编码,F.名称,F.是否变价,D.库房分批,D.在用分批,P.现价,F.规格,F.产地,D.最大效期,D.指导差价率,  " & _
                "        D.包装单位,D.换算系数,F.计算单位 "
    ElseIf optMode(根据下限) Then
       gstrSQL = "Select Nvl(A.下限,0)-Sum(Nvl(B.可用数量,0)) 申领数量,Sum(Nvl(K.可用数量,0)) 可用数量,Sum(Nvl(K.实际数量,0)) 实际数量,Sum(Nvl(K.实际金额,0)) 实际金额,Sum(Nvl(K.实际差价,0)) 实际差价,  " & _
                "         D.材料ID,F.编码,F.名称,F.是否变价,D.库房分批,D.在用分批,P.现价 售价,F.规格,F.产地,D.最大效期,D.指导差价率,  " & _
                "         D.包装单位,D.换算系数,F.计算单位 售价单位  " & _
                "  From (Select 库房id, 材料id, 上限, 下限, 盘点属性, 库房货位 From 材料储备限额 Where 库房ID=[1] And Nvl(下限,0)>0) A, " & _
                "       材料特性 D,收费项目目录 F,诊疗执行科室 Z,诊疗项目目录 N,  " & _
                "       (Select 收费细目ID,现价 From 收费价目 Where SysDate Between 执行日期 And Nvl(终止日期,Sysdate)" & _
                GetPriceClassString("") & ") P,  " & _
                "       (Select 药品id 材料ID,Sum(可用数量) as 可用数量,sum(实际数量) 实际数量, sum(实际金额) 实际金额,sum(实际差价) 实际差价 " & _
                "        From 药品库存 Where 库房ID=[1] And 性质=1 Group by 药品id) B,  " & _
                "      (Select 药品id 材料ID,Sum(可用数量) as 可用数量,sum(实际数量) 实际数量, sum(实际金额) 实际金额,sum(实际差价) 实际差价 " & _
                "      From 药品库存 Where 库房ID=[4] And 性质=1 Group by 药品id ) K " & IIf(str分类id_IN = "", "", ",Table(Cast(f_Num2List([5]) As zlTools.t_NumList)) Q") & _
                "  Where A.材料ID=D.材料ID and A.材料id=b.材料id(+)  And A.材料ID=K.材料ID(+)  And D.材料ID=P.收费细目ID" & IIf(str分类id_IN = "", "", " And N.分类id+0=Q.Column_Value") & _
                "  And D.材料ID=F.ID and D.诊疗ID=N.id  And D.诊疗ID=Z.诊疗项目ID And Z.执行科室ID+0=[1]" & _
                "  Having Nvl(A.下限,0)-Sum(Nvl(B.可用数量,0))>0 " & _
                "  Group By Nvl(A.下限,0),D.材料ID,F.编码,F.名称,F.是否变价,D.库房分批,D.在用分批,P.现价,F.规格,F.产地,D.最大效期,D.指导差价率,  " & _
                "        D.包装单位,D.换算系数,F.计算单位 "
                
    ElseIf optMode(根据申领单未发数) Then   '根据申领单未发数（不加条件And Nvl(A.发药方式,0)=1 ，是因为审核时，是删除申领单，产生移库单再审核的，标志已经没有了）
        gstrSQL = "select sum(A.填写数量-A.实际数量) 申领数量,max(Nvl(B.可用数量,0)) 可用数量,max(Nvl(B.实际数量,0)) 实际数量,max(Nvl(B.实际金额,0)) 实际金额,max(Nvl(B.实际差价,0)) 实际差价, " & _
                 "        D.材料ID,F.编码,F.名称,F.是否变价,D.库房分批,D.在用分批,P.现价 售价,F.规格,F.产地,D.最大效期,D.指导差价率, " & _
                 "        D.包装单位,D.换算系数,F.计算单位 售价单位 " & _
                 " from 药品收发记录 A,材料特性 D,收费项目目录 F,诊疗执行科室 Z,诊疗项目目录 N, " & _
                 "      (Select 收费细目ID,现价 From 收费价目 Where SysDate Between 执行日期 And Nvl(终止日期,Sysdate)" & _
                 GetPriceClassString("") & ") P, " & _
                 "      (Select 药品id 材料ID,sum(可用数量) as 可用数量,sum(实际数量) as 实际数量, sum(实际金额) as 实际金额,sum(实际差价) as 实际差价" & _
                 "      From 药品库存 Where 库房ID=[4] And 性质=1 Group by 药品id) B " & IIf(str分类id_IN = "", "", ",Table(Cast(f_Num2List([5]) As zlTools.t_NumList)) Q") & _
                 " Where A.单据=19 And A.审核日期 Between [2]" & _
                 " And [3]" & _
                 " And A.库房ID=[4] And A.对方部门ID=[1]" & _
                 " And A.药品ID=D.材料ID  and A.入出系数<>1 and D.诊疗ID=N.id  And A.药品id =b.材料id(+)" & _
                 " And D.材料ID=P.收费细目ID" & IIf(str分类id_IN = "", "", " And N.分类id+0=Q.Column_Value") & _
                 " And D.材料ID=F.ID And D.诊疗ID=Z.诊疗项目ID And Z.执行科室ID+0=[1]" & _
                 " having sum(A.填写数量-A.实际数量)>0 " & _
                 " Group By D.材料ID,F.编码,F.名称,F.是否变价,D.库房分批,D.在用分批,P.现价,F.规格,F.产地,D.最大效期,D.指导差价率, " & _
                 "       D.包装单位,D.换算系数,F.计算单位 "
    
    ElseIf optMode(根据申领单汇总) Then
        gstrSQL = "select sum(Nvl(A.实际数量,0)) 申领数量,max(Nvl(B.可用数量,0)) 可用数量,max(Nvl(B.实际数量,0)) 实际数量,max(Nvl(B.实际金额,0)) 实际金额,max(Nvl(B.实际差价,0)) 实际差价, " & _
                 "        D.材料ID,F.编码,F.名称,F.是否变价,D.库房分批,D.在用分批,P.现价 售价,F.规格,F.产地,D.最大效期,D.指导差价率, " & _
                 "        D.包装单位,D.换算系数,F.计算单位 售价单位 " & _
                 " from 药品收发记录 A,材料特性 D,收费项目目录 F,(Select Distinct 诊疗项目id, 执行科室id From 诊疗执行科室) Z,诊疗项目目录 N, " & _
                 "      (Select 收费细目ID,现价 From 收费价目 Where SysDate Between 执行日期 And Nvl(终止日期,Sysdate) " & GetPriceClassString("") & ") P, " & _
                 "      (Select 药品id 材料ID,Sum(可用数量) as 可用数量,sum(实际数量) 实际数量, sum(实际金额) 实际金额,sum(实际差价) 实际差价 " & _
                 "       From 药品库存 Where 库房ID=[4] And 性质=1 Group by 药品id) B " & IIf(str分类id_IN = "", "", ",Table(Cast(f_Num2List([5]) As zlTools.t_NumList)) Q") & _
                 " Where A.单据=19 And A.入出系数=1 And A.审核日期 Between [2]" & _
                 " And [3]" & _
                 " And A.药品ID=D.材料ID and D.诊疗ID=N.id  And Nvl(A.药品id,0)=B.材料ID(+) And A.库房ID=[1] " & _
                 " And A.药品ID=D.材料ID  And D.材料ID=P.收费细目ID" & IIf(str分类id_IN = "", "", " And N.分类id+0=Q.Column_Value") & _
                 " And D.材料ID=F.ID And D.诊疗ID=Z.诊疗项目ID And Z.执行科室ID+0=[1]" & _
                 " Having Sum(Nvl(A.实际数量,0))>0 " & _
                 " Group By D.材料ID,F.编码,F.名称,F.是否变价,D.库房分批,D.在用分批,P.现价,F.规格,F.产地,D.最大效期,D.指导差价率, " & _
                 "       D.包装单位,D.换算系数,F.计算单位 "

    End If
    
    Set rsCheck = zlDatabase.OpenSQLRecord(gstrSQL, "检查是否存在符合条件的记录", mlngStockID, CDate(Format(dtp开始时间.Value, "yyyy-MM-dd HH:mm:ss")), CDate(Format(dtp结束时间.Value, "yyyy-MM-dd HH:mm:ss")), lngTargetID, str分类id_IN)
    
    If rsCheck.RecordCount = 0 Then
        MsgBox "没找到符合条件的记录！", vbInformation, gstrSysName
        Exit Function
    End If
    
    Call WriteResult(rsCheck)
    
    Dim intCount As Integer
    With frmRequestStuffCard
        For intCount = 0 To .cboStock.ListCount - 1
            If .cboStock.ItemData(intCount) = lngTargetID Then
                .cboStock.ListIndex = intCount: Exit For
            End If
        Next
    End With
    CheckData = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub WriteResult(ByVal rsCheck As ADODB.Recordset)
    Dim strUnit As String
    Dim lngTargetID As Long
    Dim bln提示 As Boolean, bln库房 As Boolean
    Dim bln允许 As Boolean, bln特药 As Boolean       'bln允许-根据系统参数“库存检查”和用户操作来决定是否产生无库存的卫材；bln特药-当前卫材是否是时价或批次卫材
    Dim dbl申领数量 As Double, dbl填写数量 As Double
    Dim rsStock As New ADODB.Recordset  '药品库存
    Dim rsTemp  As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    lngTargetID = cbo库房.ItemData(cbo库房.ListIndex)
    Call GetPara(lngTargetID)
    bln库房 = CheckStock(lngTargetID)
    Dim blnData As Boolean
    '准备产生数据（全部以零售单位为准，最终在SetColValue函数中转换，传入的系数为当前单位的系数）
    With rsCheck
        Do While Not .EOF
            If mbln明确批次 Then
                dbl申领数量 = zlStr.NVL(!申领数量, 0)
                gstrSQL = " Select Nvl(可用数量,0) 可用数量,Nvl(实际数量,0) 实际数量,Nvl(实际金额,0) 实际金额,Nvl(实际差价,0) 实际差价," & _
                          "     Nvl(批次,0) 批次,效期,灭菌效期,上次批号 批号,上次产地 产地,批准文号 " & _
                          " From 药品库存 Where 库房ID=[1] And 药品ID=[2]  And 性质=1"
                If gSystem_Para.P156_出库算法 = 0 Then
                    gstrSQL = gstrSQL & " Order by Nvl(批次,0)"
                Else
                    gstrSQL = gstrSQL & " Order by 效期,Nvl(批次,0)"
                End If
                          
                Set rsStock = zlDatabase.OpenSQLRecord(gstrSQL, "取该卫材的库存", lngTargetID, Val(zlStr.NVL(rsCheck!材料ID)))
                          
                blnData = False
                If rsStock.RecordCount <> 0 Then
                    '有库存的卫材、分批或时价卫材按此产生
                    Do While Not rsStock.EOF
                        If dbl申领数量 >= rsStock!可用数量 Then
                            dbl填写数量 = rsStock!可用数量
                        Else
                            dbl填写数量 = dbl申领数量
                        End If
                        If rsStock!可用数量 < 0 Then
                            dbl填写数量 = 0
                        End If
                     
                         blnData = SetColValue(!材料ID, "[" & !编码 & "]" & !名称, IIf(IsNull(!规格), "", !规格), IIf(IsNull(rsStock!产地), "", rsStock!产地), _
                            IIf(mintUnit = 0, !售价单位, !包装单位), _
                            !售价, IIf(IsNull(rsStock!批号), "", rsStock!批号), _
                            IIf(IsNull(rsStock!效期), "", rsStock!效期), IIf(IsNull(!最大效期), 0, !最大效期), _
                            IIf(zlStr.NVL(rsStock!灭菌效期) = "", "", Format(rsStock!灭菌效期, "yyyy-mm-dd")), _
                            !库房分批, zlStr.NVL(!可用数量, 0), _
                            IIf(IsNull(!实际金额), 0, !实际金额), IIf(IsNull(!实际差价), 0, !实际差价), !指导差价率, _
                            IIf(mintUnit = 0, 1, !换算系数), _
                            rsStock!批次, dbl填写数量, !在用分批, !是否变价, IIf(IsNull(rsStock!批准文号), "", rsStock!批准文号))
                        
                        dbl申领数量 = dbl申领数量 - dbl填写数量
                        If dbl申领数量 = 0 Then Exit Do
                        rsStock.MoveNext
                    Loop
                    With frmRequestStuffCard.mshBill
                          If dbl申领数量 <> 0 And blnData And optMode(根据申领单汇总).Value = False Then
                            '未申领完的数量全部放在最后一行的材料上
                            .TextMatrix(.Rows - 2, mconIntCol填写数量) = Format(Val(.TextMatrix(.Rows - 2, mconIntCol填写数量)) + dbl申领数量 / IIf(Val(.TextMatrix(.Rows - 2, mconIntCol比例系数)) = 0, 1, Val(.TextMatrix(.Rows - 2, mconIntCol比例系数))), mFMT.FM_数量)
                          End If
                    End With
        
                Else
                    '不分批且无时价属性的卫材按此产生
                    '如果参数为不足禁止，根本不执行以下语句
                    If mintCheck <> 2 Then
                        gstrSQL = " Select Nvl(A.库房分批,0) 库房分批,Nvl(A.在用分批,0) 在用分批,Nvl(B.是否变价,0) 时价 " & _
                                  " From 材料特性 A,收费项目目录 B" & _
                                  " Where A.材料ID = B.ID And A.材料ID =[1] "
                        
                        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取该卫材对于出库库房是否分批、时价的属性", Val(zlStr.NVL(!材料ID)))
                                  
                        bln特药 = (rsTemp!时价 = 1) Or IIf(bln库房, (rsTemp!库房分批 = 1), (rsTemp!在用分批 = 1))
                        If Not bln特药 Then
                            If Not bln提示 Then
                                If mintCheck = 1 Then
                                    bln允许 = (MsgBox("无库存卫材是否继续申领？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes)
                                Else
                                    bln允许 = True
                                End If
                                bln提示 = True
                            End If
                            If bln允许 Then
                                '为无库存卫材产生申领记录
                                Call SetColValue(!材料ID, "[" & !编码 & "]" & !名称, IIf(IsNull(!规格), "", !规格), "", _
                                    IIf(mintUnit = 0, !售价单位, !包装单位), _
                                    !售价, "", "", IIf(IsNull(!最大效期), 0, !最大效期), "", !库房分批, IIf(IsNull(!可用数量), 0, !可用数量), _
                                    IIf(IsNull(!实际金额), 0, !实际金额), IIf(IsNull(!实际差价), 0, !实际差价), !指导差价率, _
                                    IIf(mintUnit = 0, 1, !换算系数), _
                                    0, !申领数量, !在用分批, !是否变价, "")
                            End If
                        End If
                    End If
                End If
            Else
                '依据传入记录集产生数据
                Call SetColValue(!材料ID, "[" & !编码 & "]" & !名称, IIf(IsNull(!规格), "", !规格), IIf(IsNull(!产地), "", !产地), _
                    IIf(mintUnit = 0, !售价单位, !包装单位), _
                    !售价, "", "", IIf(IsNull(!最大效期), 0, !最大效期), "", !库房分批, IIf(IsNull(!可用数量), 0, !可用数量), _
                    IIf(IsNull(!实际金额), 0, !实际金额), IIf(IsNull(!实际差价), 0, !实际差价), !指导差价率, _
                    IIf(mintUnit = 0, 1, !换算系数), _
                    0, !申领数量, !在用分批, !是否变价, "")
            End If
            .MoveNext
        Loop
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'从卫材目录中取值并附给相应的列
Private Function SetColValue(ByVal lng材料ID As Long, ByVal str药名 As String, ByVal str规格 As String, _
    ByVal str产地 As String, ByVal str单位 As String, ByVal num售价 As Double, ByVal str批号 As String, _
    ByVal str效期 As String, ByVal int最大效期 As Integer, ByVal str灭菌失效期 As String, ByVal int分批核算 As Integer, _
    ByVal num可用数量 As Double, ByVal num实际金额 As Double, ByVal num实际差价 As Double, _
    ByVal num指导差价率 As Double, ByVal num比例系数 As Double, ByVal lng批次 As Long, _
    ByVal dbl数量 As Double, ByVal int在用分批 As Integer, ByVal int是否变价 As Integer, ByVal str批准文号 As String) As Boolean
    
      
    Dim intCount As Integer
    Dim intCol As Integer
    Dim intRow As Integer
    
    
    Dim num实际数量 As Double
    Dim rsTemp As New ADODB.Recordset
        
 
    On Error GoTo ErrHandle
       
    SetColValue = False
    
    '如果申领数量为零则退出
    If IIf(dbl数量 >= num可用数量, num可用数量, dbl数量) = 0 And mbln明确批次 And (int是否变价 = 1 Or lng批次 <> 0) Then Exit Function
    
    '如果是时价材料且明确批次,需要重新计算售价;否则以传入零售价为准
    If int是否变价 = 1 Then
        '取实际数量
        gstrSQL = " Select nvl(实际金额,0) 实际金额,Nvl(实际数量,0) 实际数量 From 药品库存 " & _
                " Where 性质=1 And 药品ID=[1] And 库房ID=[2] And Nvl(批次,0)=[3]"
        
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "如果是时价材料,需要重新计算售价", lng材料ID, cbo库房.ItemData(cbo库房.ListIndex), lng批次)
        If Not rsTemp.EOF Then
            If rsTemp!实际数量 > 0 Then
                num售价 = rsTemp!实际金额 / rsTemp!实际数量
            End If
        End If
    End If
    
    With frmRequestStuffCard.mshBill
        intRow = .Rows - 1
        .TextMatrix(intRow, 0) = lng材料ID
        .TextMatrix(intRow, 1) = intRow
        .TextMatrix(intRow, mconIntCol药名) = str药名
        .TextMatrix(intRow, mconIntCol规格) = str规格
        .TextMatrix(intRow, mconIntCol产地) = str产地
        .TextMatrix(intRow, mconIntCol批准文号) = str批准文号
        .TextMatrix(intRow, mconIntCol单位) = str单位
        .TextMatrix(intRow, mconIntCol批号) = str批号
        .TextMatrix(intRow, mconIntCol效期) = Format(str效期, "yyyy-mm-dd")
        .TextMatrix(intRow, mconIntCol灭菌失效期) = Format(str灭菌失效期, "yyyy-mm-dd")
        
        .TextMatrix(intRow, mconIntCol售价) = Format(num售价 * num比例系数, mFMT.FM_零售价)
        .TextMatrix(intRow, mconIntCol分批核算) = int分批核算
        .TextMatrix(intRow, mconIntCol可用数量) = Format(num可用数量 / num比例系数, mFMT.FM_数量)
        .TextMatrix(intRow, mconIntCol最大效期) = int最大效期 & "||" & int是否变价 & "||" & int在用分批
        .TextMatrix(intRow, mconIntCol实际差价) = num实际差价
        .TextMatrix(intRow, mconIntCol实际金额) = num实际金额
        .TextMatrix(intRow, mconIntCol指导差价率) = num指导差价率
        .TextMatrix(intRow, mconIntCol比例系数) = num比例系数
        .TextMatrix(intRow, mconIntCol批次) = lng批次
        '如果是时价材料或分批材料,不能超过当前库存数量
        If (int是否变价 = 1 Or lng批次 <> 0) And mbln明确批次 Then
            .TextMatrix(intRow, mconIntCol填写数量) = Format(IIf(dbl数量 >= num可用数量, num可用数量, dbl数量) / num比例系数, mFMT.FM_数量)
            .TextMatrix(intRow, mconIntCol实际数量) = Format(IIf(dbl数量 >= num可用数量, num可用数量, dbl数量) / num比例系数, mFMT.FM_数量)
        Else
            .TextMatrix(intRow, mconIntCol填写数量) = Format(dbl数量 / num比例系数, mFMT.FM_数量)
            .TextMatrix(intRow, mconIntCol实际数量) = Format(dbl数量 / num比例系数, mFMT.FM_数量)
        End If
        
        If .TextMatrix(intRow, mconIntCol售价) <> "" Then
            .TextMatrix(intRow, mconIntCol售价金额) = Format(.TextMatrix(intRow, mconIntCol售价) * .TextMatrix(intRow, mconIntCol填写数量), mFMT.FM_金额)
        End If
        
        
        Dim dbl差价 As Double, dbl购价 As Double, dbl成本金额 As Double
        'cboStock.ItemData(cboStock.ListIndex)
        Call 验证出库差价计算(cbo库房.ItemData(cbo库房.ListIndex), lng材料ID, lng批次, _
            num比例系数 + 0, num实际差价, num实际金额, _
            num指导差价率 / 100, Val(.TextMatrix(intRow, mconIntCol填写数量)), Val(.TextMatrix(intRow, mconIntCol售价金额)), _
            dbl差价, dbl购价, dbl成本金额)
            
        .TextMatrix(intRow, mconintCol差价) = Format(dbl差价, mFMT.FM_金额)
        .TextMatrix(intRow, mconIntCol采购价) = Format(dbl购价, mFMT.FM_成本价)
        .TextMatrix(intRow, mconIntCol采购金额) = Format(dbl成本金额, mFMT.FM_金额)
'
'
'        If .TextMatrix(intRow, mconIntCol实际金额) = 0 Then
'            .TextMatrix(intRow, mconintCol差价) = Format(.TextMatrix(intRow, mconIntCol售价金额) * .TextMatrix(intRow, mconIntCol指导差价率) / 100, mFMT.FM_金额)
'        Else
'            .TextMatrix(intRow, mconintCol差价) = Format(.TextMatrix(intRow, mconIntCol售价金额) * (.TextMatrix(intRow, mconIntCol实际差价) / .TextMatrix(intRow, mconIntCol实际金额)), mFMT.FM_金额)
'        End If
'        .TextMatrix(intRow, mconIntCol采购价) = Format((.TextMatrix(intRow, mconIntCol售价金额) - .TextMatrix(intRow, mconintCol差价)) / IIf(Val(.TextMatrix(intRow, mconIntCol填写数量)) = 0, 1, Val(.TextMatrix(intRow, mconIntCol填写数量))), mFMT.FM_成本价)
'        .TextMatrix(intRow, mconIntCol采购金额) = Format(.TextMatrix(intRow, mconIntCol采购价) * .TextMatrix(intRow, mconIntCol填写数量), mFMT.FM_金额)
'
        .Rows = .Rows + 1
    End With
    SetColValue = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckStock(ByVal lng库房ID As Long) As Boolean
    Dim rsCheck As New ADODB.Recordset
    '检查指定库房是卫材库、发料部门还是制剂室(传入的库房肯定是卫材库、发料部门或制剂室中的一个)
    
    On Error GoTo ErrHandle
    gstrSQL = " Select 部门ID From 部门性质说明 " & _
              " Where (工作性质 like '发料部门' Or 工作性质 like '%制剂室') And 部门id=[1]"
    Set rsCheck = zlDatabase.OpenSQLRecord(gstrSQL, "判断是不是发料部门或制剂室", lng库房ID)
              
    If rsCheck.EOF Then
        CheckStock = True
    End If
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub GetPara(ByVal lng库房ID As Long)
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo ErrHandle
    '获取出库检查的参数设置值（0-不检查;1-检查，不足提醒;2-不足禁止）
    gstrSQL = " Select Nvl(检查方式,0) Value From 材料出库检查 Where 库房ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取检查库存的参数", lng库房ID)
    
    If Not rsTemp.EOF Then
        mintCheck = rsTemp!Value
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub




Private Sub SetParentNode(ByVal Node As MSComctlLib.Node, blnCheck As Boolean)
    Dim intIdx As Integer
    
    If Not Node.Parent Is Nothing Then
        If blnCheck = True Then
            '看是否他的兄弟接点是否也全是TRUE，如是，则置其父节点也为TRUE，否则，不管
            intIdx = Node.FirstSibling.Index
            Do While intIdx <> Node.LastSibling.Index
                If tvw分类.Nodes(intIdx).Checked = False Then
                    Node.Parent.Checked = False
                    Exit Do
                End If
                intIdx = tvw分类.Nodes(intIdx).Next.Index
            Loop
            If intIdx = Node.LastSibling.Index Then
                If tvw分类.Nodes(intIdx).Checked = True Then
                    Node.Parent.Checked = True
                End If
            End If
        Else
            Node.Parent.Checked = False
        End If
        
        Set Node = Node.Parent
        If Not Node Is Nothing Then
            SetParentNode Node, blnCheck
        End If
    End If
End Sub


Private Function CheckNode(ByVal Node As Object, blnCheck As Boolean)
    Dim intIdx As Integer
    
    If Node.Children > 0 Then
        Set Node = Node.Child
        Do While Not Node Is Nothing
            Node.Checked = blnCheck
            If Node.Children > 0 Then
                CheckNode Node, blnCheck
            End If
            Set Node = Node.Next
        Loop
    Else
        Node.Checked = blnCheck
    End If
End Function

Private Function CheckCount() As Integer
    Dim i As Integer
    For i = 1 To tvw分类.Nodes.count
        If tvw分类.Nodes(i).Checked Then CheckCount = CheckCount + 1
    Next
End Function

Private Sub tvw分类_NodeCheck(ByVal Node As MSComctlLib.Node)
    CheckNode Node, Node.Checked
    SetParentNode Node, Node.Checked
End Sub



