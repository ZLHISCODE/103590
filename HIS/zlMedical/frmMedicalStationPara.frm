VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmMedicalStationPara 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   5760
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5925
   Icon            =   "frmMedicalStationPara.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   150
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   5265
      Width           =   1100
   End
   Begin TabDlg.SSTab tbs 
      Height          =   5085
      Left            =   60
      TabIndex        =   38
      Top             =   60
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   8969
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "&1.体检"
      TabPicture(0)   =   "frmMedicalStationPara.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame6"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame7"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "&2.记费"
      TabPicture(1)   =   "frmMedicalStationPara.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(1)=   "fraAction"
      Tab(1).Control(2)=   "Frame2"
      Tab(1).Control(3)=   "Frame3"
      Tab(1).Control(4)=   "Frame4"
      Tab(1).Control(5)=   "lst收费类别"
      Tab(1).Control(6)=   "Label1"
      Tab(1).ControlCount=   7
      Begin VB.Frame Frame7 
         Caption         =   "正在体检范围"
         Height          =   1095
         Left            =   120
         TabIndex        =   47
         Top             =   2040
         Width           =   5475
         Begin VB.CheckBox chk 
            Caption         =   "按报到时间查(&8)"
            Height          =   240
            Index           =   2
            Left            =   3795
            TabIndex        =   50
            Top             =   750
            Width           =   1650
         End
         Begin VB.ComboBox cbo 
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   5
            Left            =   1875
            Style           =   2  'Dropdown List
            TabIndex        =   49
            Top             =   690
            Width           =   1920
         End
         Begin VB.ComboBox cbo 
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   1
            Left            =   1875
            Style           =   2  'Dropdown List
            TabIndex        =   48
            Top             =   330
            Width           =   1920
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "正体检团体时间(&7)"
            Height          =   180
            Index           =   6
            Left            =   285
            TabIndex        =   52
            Top             =   750
            Width           =   1530
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "正体检个人时间(&6)"
            Height          =   180
            Index           =   2
            Left            =   285
            TabIndex        =   51
            Top             =   375
            Width           =   1530
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "其他"
         Height          =   1695
         Left            =   120
         TabIndex        =   46
         Top             =   3225
         Width           =   5475
         Begin VB.CheckBox chk 
            Caption         =   "体检人员接受或报到时自动打印指引单(&D)"
            Height          =   270
            Index           =   1
            Left            =   315
            TabIndex        =   12
            Top             =   1125
            Width           =   3810
         End
         Begin VB.CheckBox chk 
            Caption         =   "保存体检结果报告时提示检查小结(&A)"
            Height          =   225
            Index           =   0
            Left            =   330
            TabIndex        =   8
            Top             =   540
            Width           =   3375
         End
         Begin VB.TextBox txt 
            Alignment       =   2  'Center
            Height          =   300
            Index           =   1
            Left            =   3135
            Locked          =   -1  'True
            TabIndex        =   10
            Text            =   "5"
            Top             =   780
            Width           =   345
         End
         Begin VB.ComboBox cbo 
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   4
            Left            =   1845
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   210
            Width           =   1920
         End
         Begin MSComCtl2.UpDown udn 
            Height          =   300
            Index           =   1
            Left            =   3480
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   780
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Value           =   5
            BuddyControl    =   "txt(1)"
            BuddyDispid     =   196616
            BuddyIndex      =   1
            OrigLeft        =   4320
            OrigTop         =   1065
            OrigRight       =   4560
            OrigBottom      =   1365
            Max             =   30
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.CheckBox chk 
            Caption         =   "体检人员接受或报到时自动打印申请单(&E)"
            Height          =   255
            Index           =   3
            Left            =   300
            TabIndex        =   13
            Top             =   1410
            Width           =   3810
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "待体检、正在体检自动刷新间隔(&B)         分"
            Height          =   180
            Index           =   4
            Left            =   300
            TabIndex        =   9
            Top             =   840
            Width           =   3780
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "体检费用费别(&9)"
            Height          =   180
            Index           =   5
            Left            =   330
            TabIndex        =   6
            Top             =   285
            Width           =   1350
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "其他"
         Height          =   1440
         Left            =   -71685
         TabIndex        =   44
         Top             =   3540
         Width           =   2280
         Begin VB.CheckBox chkPay 
            Caption         =   "中药可以输入付数(&P)"
            Height          =   195
            Left            =   150
            TabIndex        =   33
            Top             =   540
            Value           =   1  'Checked
            Width           =   2055
         End
         Begin VB.CheckBox chkTime 
            Caption         =   "变价允许输入数次(&N)"
            Height          =   195
            Left            =   150
            TabIndex        =   32
            Top             =   255
            Width           =   2055
         End
         Begin VB.CheckBox chk药库 
            Caption         =   "显示其它药库库存(&T)"
            Height          =   195
            Left            =   150
            TabIndex        =   35
            Top             =   1095
            Width           =   2055
         End
         Begin VB.CheckBox chk药房 
            Caption         =   "显示其它药房库存(&R)"
            Height          =   195
            Left            =   150
            TabIndex        =   34
            Top             =   810
            Width           =   2055
         End
      End
      Begin VB.Frame fraAction 
         Caption         =   "执行设置 "
         Height          =   840
         Left            =   -74865
         TabIndex        =   43
         Top             =   4140
         Width           =   3135
         Begin VB.CheckBox chkFinish 
            Caption         =   "允许完成未收费病人的项目(&L)"
            Height          =   195
            Left            =   120
            TabIndex        =   29
            Top             =   540
            Width           =   2805
         End
         Begin VB.CheckBox chkActLog 
            Caption         =   "允许他人代行执行记录(&K)"
            Height          =   195
            Left            =   120
            TabIndex        =   28
            Top             =   285
            Width           =   2745
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "住院药房设置 "
         Height          =   1305
         Left            =   -74880
         TabIndex        =   42
         Top             =   1830
         Width           =   3135
         Begin VB.ComboBox cbo住中药 
            Height          =   300
            Left            =   975
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   930
            Width           =   1950
         End
         Begin VB.ComboBox cbo住西药 
            Height          =   300
            Left            =   975
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   240
            Width           =   1950
         End
         Begin VB.ComboBox cbo住成药 
            Height          =   300
            Left            =   975
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   585
            Width           =   1950
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "中草药(&G)"
            Height          =   180
            Left            =   120
            TabIndex        =   24
            Top             =   990
            Width           =   810
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "西成药(&E)"
            Height          =   180
            Left            =   120
            TabIndex        =   20
            Top             =   300
            Width           =   810
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "中成药(&F)"
            Height          =   180
            Left            =   120
            TabIndex        =   22
            Top             =   645
            Width           =   810
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "门诊药房设置 "
         Height          =   1320
         Left            =   -74880
         TabIndex        =   41
         Top             =   480
         Width           =   3135
         Begin VB.ComboBox cbo门成药 
            Height          =   300
            Left            =   975
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   600
            Width           =   1950
         End
         Begin VB.ComboBox cbo门西药 
            Height          =   300
            Left            =   975
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   255
            Width           =   1950
         End
         Begin VB.ComboBox cbo门中药 
            Height          =   300
            Left            =   975
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   945
            Width           =   1950
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "中成药(&B)"
            Height          =   180
            Left            =   105
            TabIndex        =   16
            Top             =   660
            Width           =   810
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "西成药(&A)"
            Height          =   180
            Left            =   105
            TabIndex        =   14
            Top             =   315
            Width           =   810
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "中草药(&D)"
            Height          =   180
            Left            =   105
            TabIndex        =   18
            Top             =   1005
            Width           =   810
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "药品单位 "
         Height          =   900
         Left            =   -74865
         TabIndex        =   40
         Top             =   3195
         Width           =   3135
         Begin VB.OptionButton opt药品单位 
            Caption         =   "售价单位(I)"
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   0
            Left            =   150
            TabIndex        =   26
            Top             =   315
            Value           =   -1  'True
            Width           =   1500
         End
         Begin VB.OptionButton opt药品单位 
            Caption         =   "门诊/住院单位(&J)"
            ForeColor       =   &H00800000&
            Height          =   180
            Index           =   1
            Left            =   150
            TabIndex        =   27
            Top             =   600
            Width           =   1845
         End
      End
      Begin VB.ListBox lst收费类别 
         Height          =   2790
         Left            =   -71670
         Style           =   1  'Checkbox
         TabIndex        =   31
         ToolTipText     =   "请复选允许使用的收费类别"
         Top             =   660
         Width           =   2265
      End
      Begin VB.Frame Frame5 
         Caption         =   "时间范围"
         Height          =   1440
         Left            =   120
         TabIndex        =   39
         Top             =   570
         Width           =   5475
         Begin VB.ComboBox cbo 
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   3
            Left            =   1830
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   1035
            Width           =   1920
         End
         Begin VB.ComboBox cbo 
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   2
            Left            =   1830
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   675
            Width           =   1920
         End
         Begin VB.ComboBox cbo 
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   0
            Left            =   1830
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   330
            Width           =   1920
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "历次体检范围(&5)"
            Height          =   180
            Index           =   0
            Left            =   240
            TabIndex        =   4
            Top             =   1110
            Width           =   1350
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "已体检缺省范围(&4)"
            Height          =   180
            Index           =   3
            Left            =   240
            TabIndex        =   2
            Top             =   750
            Width           =   1530
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "待体检缺省范围(&3)"
            Height          =   180
            Index           =   1
            Left            =   240
            TabIndex        =   0
            Top             =   405
            Width           =   1530
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "输入类别(&M)"
         Height          =   180
         Left            =   -71685
         TabIndex        =   30
         Top             =   435
         Width           =   990
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3480
      TabIndex        =   36
      Top             =   5265
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4650
      TabIndex        =   37
      Top             =   5265
      Width           =   1100
   End
End
Attribute VB_Name = "frmMedicalStationPara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mblnOK As Boolean
Private mlngLoop As Long
Private mfrmMain As Object

Public Function ShowPara(ByVal frmMain As Object) As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim objCbo As ComboBox, lng药房ID As Long
    Dim strSQL As String, strPar As String, i As Long
    Dim rs As New ADODB.Recordset
    
    mblnOK = False
    
    Set mfrmMain = frmMain
    '初始化
    
    For mlngLoop = 0 To 5
        If mlngLoop <> 4 Then
            cbo(mlngLoop).AddItem "今  天"
            cbo(mlngLoop).AddItem "昨  天"
            cbo(mlngLoop).AddItem "本  周"
            cbo(mlngLoop).AddItem "本  月"
            cbo(mlngLoop).AddItem "本  季"
            cbo(mlngLoop).AddItem "本半年"
            cbo(mlngLoop).AddItem "本  年"
            cbo(mlngLoop).AddItem "前三天"
            cbo(mlngLoop).AddItem "前一周"
            cbo(mlngLoop).AddItem "前半月"
            cbo(mlngLoop).AddItem "前一月"
            cbo(mlngLoop).AddItem "前二月"
            cbo(mlngLoop).AddItem "前三月"
            cbo(mlngLoop).AddItem "前半年"
            cbo(mlngLoop).AddItem "前一年"
            cbo(mlngLoop).AddItem "前二年"
        End If
    Next
    
    cbo(4).Clear
    cbo(4).AddItem ""
    gstrSQL = "Select 名称,1 As ID From 费别"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If rs.BOF = False Then
        Call AddComboData(cbo(4), rs, False)
    End If
    
    On Error Resume Next
    
    '保存的是体检费用的费别名称
    
    cbo(4).Text = Trim(zlDatabase.GetPara(130, glngSys, , ""))
    chk(0).Value = Val(zlDatabase.GetPara(131, glngSys, , "0"))
    
    cbo(0).Text = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & mfrmMain.Name, "待体检时间范围", "今  天")
    cbo(1).Text = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & mfrmMain.Name, "正体检时间范围", "今  天")
    
    cbo(5).Text = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & mfrmMain.Name, "正体检团体时间范围", "今  天")
    chk(2).Value = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & mfrmMain.Name, "正体检查询依据", "0"))
    
    cbo(2).Text = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & mfrmMain.Name, "已完体检时间范围", "今  天")
    cbo(3).Text = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & mfrmMain.Name, "历次体检范围", "今  天")
    txt(1).Text = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & mfrmMain.Name, "自动刷新间隔", 5))
    chk(1).Value = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & mfrmMain.Name, "自动打印指引单", 0))
    chk(3).Value = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & mfrmMain.Name, "自动打印申请单", 0))
    
    If cbo(0).ListIndex = -1 Then cbo(0).ListIndex = 0
    If cbo(1).ListIndex = -1 Then cbo(1).ListIndex = 0
    If cbo(2).ListIndex = -1 Then cbo(2).ListIndex = 0
    
    
    chkPay.Value = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\zl9CISWork", "中药付数", 1))
    chkTime.Value = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\zl9CISWork", "变价数次", 0))
    chk药房.Value = Val(GetSetting("ZLSOFT", "公共模块\zl9CISWork", "显示其它药房库存", 0))
    chk药库.Value = Val(GetSetting("ZLSOFT", "公共模块\zl9CISWork", "显示其它药库库存", 0))
    
    '药品单位
    i = Val(GetSetting("ZLSOFT", "公共模块\zl9CISWork", "药品单位", 0))
    opt药品单位(IIf(i = 0, 0, 1)).Value = True
    
    '缺省药房
    cbo门西药.AddItem "系统分配": cbo门西药.ListIndex = 0
    cbo门成药.AddItem "系统分配": cbo门成药.ListIndex = 0
    cbo门中药.AddItem "系统分配": cbo门中药.ListIndex = 0
    cbo住西药.AddItem "系统分配": cbo住西药.ListIndex = 0
    cbo住成药.AddItem "系统分配": cbo住成药.ListIndex = 0
    cbo住中药.AddItem "系统分配": cbo住中药.ListIndex = 0
    strSQL = _
        "Select Distinct A.ID,A.编码,A.名称,B.工作性质,B.服务对象" & _
        " From 部门表 A,部门性质说明 B " & _
        " Where (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
        " And B.部门ID=A.ID And B.服务对象 IN(1,2,3)" & _
        " And B.工作性质 in('西药房','成药房','中药房')" & _
        " Order by A.编码"
        
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    For i = 1 To rsTmp.RecordCount
        If rsTmp!工作性质 = "西药房" Then
            Set objCbo = IIf(rsTmp!服务对象 = 1, cbo门西药, IIf(rsTmp!服务对象 = 2, cbo住西药, Nothing))
        End If
        If rsTmp!工作性质 = "成药房" Then
            Set objCbo = IIf(rsTmp!服务对象 = 1, cbo门成药, IIf(rsTmp!服务对象 = 2, cbo住成药, Nothing))
        End If
        If rsTmp!工作性质 = "中药房" Then
            Set objCbo = IIf(rsTmp!服务对象 = 1, cbo门中药, IIf(rsTmp!服务对象 = 2, cbo住中药, Nothing))
        End If
        If objCbo Is Nothing Then
            If rsTmp!工作性质 = "西药房" Then
                cbo门西药.AddItem rsTmp!名称
                cbo门西药.ItemData(cbo门西药.NewIndex) = rsTmp!ID
                cbo住西药.AddItem rsTmp!名称
                cbo住西药.ItemData(cbo住西药.NewIndex) = rsTmp!ID
            ElseIf rsTmp!工作性质 = "成药房" Then
                cbo门成药.AddItem rsTmp!名称
                cbo门成药.ItemData(cbo门成药.NewIndex) = rsTmp!ID
                cbo住成药.AddItem rsTmp!名称
                cbo住成药.ItemData(cbo住成药.NewIndex) = rsTmp!ID
            ElseIf rsTmp!工作性质 = "中药房" Then
                cbo门中药.AddItem rsTmp!名称
                cbo门中药.ItemData(cbo门中药.NewIndex) = rsTmp!ID
                cbo住中药.AddItem rsTmp!名称
                cbo住中药.ItemData(cbo住中药.NewIndex) = rsTmp!ID
            End If
        Else
            objCbo.AddItem rsTmp!名称
            objCbo.ItemData(objCbo.NewIndex) = rsTmp!ID
        End If
        rsTmp.MoveNext
    Next
    lng药房ID = Val(GetSetting("ZLSOFT", "公共模块\zl9CISWork", "门诊缺省西药房", 0))
    Call zlControl.CboLocate(cbo门西药, lng药房ID, True)
    lng药房ID = Val(GetSetting("ZLSOFT", "公共模块\zl9CISWork", "门诊缺省成药房", 0))
    Call zlControl.CboLocate(cbo门成药, lng药房ID, True)
    lng药房ID = Val(GetSetting("ZLSOFT", "公共模块\zl9CISWork", "门诊缺省中药房", 0))
    Call zlControl.CboLocate(cbo门中药, lng药房ID, True)
    lng药房ID = Val(GetSetting("ZLSOFT", "公共模块\zl9CISWork", "住院缺省西药房", 0))
    Call zlControl.CboLocate(cbo住西药, lng药房ID, True)
    lng药房ID = Val(GetSetting("ZLSOFT", "公共模块\zl9CISWork", "住院缺省成药房", 0))
    Call zlControl.CboLocate(cbo住成药, lng药房ID, True)
    lng药房ID = Val(GetSetting("ZLSOFT", "公共模块\zl9CISWork", "住院缺省中药房", 0))
    Call zlControl.CboLocate(cbo住中药, lng药房ID, True)
    
    '收费类别
    strSQL = "Select 编码,名称 as 类别 From 收费项目类别 Where 编码<>'1' Order by 编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Do While Not rsTmp.EOF
        lst收费类别.AddItem rsTmp!类别
        lst收费类别.ItemData(lst收费类别.NewIndex) = Asc(rsTmp!编码)
        rsTmp.MoveNext
    Loop
    strPar = GetSetting("ZLSOFT", "公共模块\zl9CISWork", "收费类别", "")
    If strPar = "" Then
        For i = 0 To lst收费类别.ListCount - 1
            lst收费类别.Selected(i) = True
        Next
    Else
        For i = 0 To lst收费类别.ListCount - 1
            If InStr(strPar, Chr(lst收费类别.ItemData(i))) Then lst收费类别.Selected(i) = True
        Next
    End If
    If lst收费类别.ListCount > 0 Then lst收费类别.TopIndex = 0: lst收费类别.ListIndex = 0
    
    '是否允许代行执行记录
    chkActLog.Value = Val(GetSetting("ZLSOFT", "公共模块\zl9CISWork", "代行执行记录", 0))
    
    '是否允许完成未收费病人的项目
    chkFinish.Value = Val(GetSetting("ZLSOFT", "公共模块\zl9CISWork", "未收费完成", 0))

    Me.Show 1, frmMain
    
    ShowPara = mblnOK
    
End Function

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cbo门成药_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cbo门西药_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cbo门中药_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cbo住成药_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cbo住西药_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cbo住中药_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chkActLog_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chkFinish_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chkPay_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chkTime_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chk药房_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chk药库_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdOK_Click()
    Dim strPar As String, i As Long
    
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & mfrmMain.Name, "待体检时间范围", cbo(0).Text)
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & mfrmMain.Name, "正体检时间范围", cbo(1).Text)
    
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & mfrmMain.Name, "正体检团体时间范围", cbo(5).Text)
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & mfrmMain.Name, "正体检查询依据", chk(2).Value)

    
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & mfrmMain.Name, "已完体检时间范围", cbo(2).Text)
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & mfrmMain.Name, "历次体检范围", cbo(3).Text)
    
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & mfrmMain.Name, "自动刷新间隔", Val(txt(1).Text))
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & mfrmMain.Name, "自动打印指引单", chk(1).Value)
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & mfrmMain.Name, "自动打印申请单", chk(3).Value)
    
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\zl9CISWork", "中药付数", chkPay.Value
    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\zl9CISWork", "变价数次", chkTime.Value
    SaveSetting "ZLSOFT", "公共模块\zl9CISWork", "显示其它药房库存", chk药房.Value
    SaveSetting "ZLSOFT", "公共模块\zl9CISWork", "显示其它药库库存", chk药库.Value
    
    '药品单位
    SaveSetting "ZLSOFT", "公共模块\zl9CISWork", "药品单位", IIf(opt药品单位(0).Value, 0, 1)
    
    '缺省药房
    SaveSetting "ZLSOFT", "公共模块\zl9CISWork", "门诊缺省西药房", cbo门西药.ItemData(cbo门西药.ListIndex)
    SaveSetting "ZLSOFT", "公共模块\zl9CISWork", "门诊缺省成药房", cbo门成药.ItemData(cbo门成药.ListIndex)
    SaveSetting "ZLSOFT", "公共模块\zl9CISWork", "门诊缺省中药房", cbo门中药.ItemData(cbo门中药.ListIndex)
    
    SaveSetting "ZLSOFT", "公共模块\zl9CISWork", "住院缺省西药房", cbo住西药.ItemData(cbo住西药.ListIndex)
    SaveSetting "ZLSOFT", "公共模块\zl9CISWork", "住院缺省成药房", cbo住成药.ItemData(cbo住成药.ListIndex)
    SaveSetting "ZLSOFT", "公共模块\zl9CISWork", "住院缺省中药房", cbo住中药.ItemData(cbo住中药.ListIndex)
    
    '收费类别
    For i = lst收费类别.ListCount - 1 To 0 Step -1
        If lst收费类别.Selected(i) Then strPar = strPar & "'" & Chr(lst收费类别.ItemData(i)) & "',"
    Next
    If strPar <> "" Then strPar = Left(strPar, Len(strPar) - 1)
    SaveSetting "ZLSOFT", "公共模块\zl9CISWork", "收费类别", strPar
    
    '是否允许代行执行记录
    SaveSetting "ZLSOFT", "公共模块\zl9CISWork", "代行执行记录", chkActLog.Value

    '是否允许完成未收费病人的项目
    SaveSetting "ZLSOFT", "公共模块\zl9CISWork", "未收费完成", chkFinish.Value
    
    Call zlDatabase.SetPara(130, cbo(4).Text, glngSys)
    Call zlDatabase.SetPara(131, chk(0).Value, glngSys)
    
    mblnOK = True

    
    Unload Me
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub lst收费类别_ItemCheck(Item As Integer)
    If lst收费类别.SelCount = 0 And Not lst收费类别.Selected(Item) Then
        lst收费类别.Selected(Item) = True
    End If
End Sub

Private Sub lst收费类别_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub opt药品单位_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub tbs_Click(PreviousTab As Integer)
    tbs.ZOrder 0
End Sub

Private Sub txt_GotFocus(Index As Integer)
    zlControl.TxtSelAll txt(Index)
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        tbs.Tab = 1
        cbo门西药.SetFocus
    End If
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not StrIsValid(txt(Index).Text, txt(Index).MaxLength)
End Sub


