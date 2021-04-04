VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOpsStationPara 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   5760
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5925
   Icon            =   "frmOpsStationPara.frx":0000
   LinkTopic       =   "Form1"
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
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   5265
      Width           =   1100
   End
   Begin TabDlg.SSTab tbs 
      Height          =   5085
      Left            =   60
      TabIndex        =   6
      Top             =   60
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   8969
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "&1.手术 "
      TabPicture(0)   =   "frmOpsStationPara.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl(8)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fra"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txt"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "udn"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdDeviceSetup"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "chk"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      Begin VB.CheckBox chk 
         Caption         =   "手术医嘱发送时不自动生成费用"
         Height          =   255
         Left            =   135
         TabIndex        =   31
         Top             =   4320
         Width           =   3045
      End
      Begin VB.CommandButton cmdDeviceSetup 
         Caption         =   "设备配置(&S)"
         Height          =   350
         Left            =   4065
         TabIndex        =   30
         Top             =   4650
         Width           =   1500
      End
      Begin MSComCtl2.UpDown udn 
         Height          =   300
         Left            =   1620
         TabIndex        =   24
         Top             =   4665
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         BuddyControl    =   "txt"
         BuddyDispid     =   196613
         OrigLeft        =   2205
         OrigTop         =   4635
         OrigRight       =   2460
         OrigBottom      =   4965
         Max             =   60
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txt 
         Height          =   300
         Left            =   1125
         MaxLength       =   2
         TabIndex        =   23
         Top             =   4665
         Width           =   495
      End
      Begin VB.Frame Frame1 
         Caption         =   "缺省执行科室"
         Height          =   2175
         Left            =   120
         TabIndex        =   10
         Top             =   2085
         Width           =   5475
         Begin VB.ComboBox cbo 
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   9
            Left            =   3750
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   1770
            Width           =   1620
         End
         Begin VB.ComboBox cbo 
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   8
            Left            =   3750
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   1410
            Width           =   1620
         End
         Begin VB.ComboBox cbo 
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   7
            Left            =   3750
            Style           =   2  'Dropdown List
            TabIndex        =   22
            Top             =   1020
            Width           =   1620
         End
         Begin VB.ComboBox cbo 
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   6
            Left            =   3750
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   645
            Width           =   1620
         End
         Begin VB.ComboBox cbo 
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   5
            Left            =   3750
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   270
            Width           =   1620
         End
         Begin VB.ComboBox cbo 
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   4
            Left            =   990
            Style           =   2  'Dropdown List
            TabIndex        =   19
            Top             =   1020
            Width           =   1620
         End
         Begin VB.ComboBox cbo 
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   3
            Left            =   990
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   645
            Width           =   1620
         End
         Begin VB.ComboBox cbo 
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   2
            Left            =   990
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   270
            Width           =   1620
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "治疗"
            Height          =   180
            Index           =   10
            Left            =   3255
            TabIndex        =   29
            Top             =   1830
            Width           =   360
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "材料"
            Height          =   180
            Index           =   9
            Left            =   3255
            TabIndex        =   27
            Top             =   1470
            Width           =   360
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "住院中药"
            Height          =   180
            Index           =   7
            Left            =   2880
            TabIndex        =   16
            Top             =   1095
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "住院成药"
            Height          =   180
            Index           =   6
            Left            =   2880
            TabIndex        =   15
            Top             =   750
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "住院西药"
            Height          =   180
            Index           =   5
            Left            =   2880
            TabIndex        =   14
            Top             =   375
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "门诊中药"
            Height          =   180
            Index           =   4
            Left            =   165
            TabIndex        =   13
            Top             =   1080
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "门诊成药"
            Height          =   180
            Index           =   2
            Left            =   180
            TabIndex        =   12
            Top             =   705
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "门诊西药"
            Height          =   180
            Index           =   0
            Left            =   165
            TabIndex        =   11
            Top             =   315
            Width           =   720
         End
      End
      Begin VB.Frame fra 
         Caption         =   "缺省时间"
         Height          =   1560
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   5475
         Begin VB.ComboBox cbo 
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   1
            Left            =   2460
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   1155
            Width           =   1920
         End
         Begin VB.ComboBox cbo 
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   0
            Left            =   2460
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   810
            Width           =   1920
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "已完成手术范围(&6)"
            Height          =   180
            Index           =   3
            Left            =   870
            TabIndex        =   2
            Top             =   1215
            Width           =   1530
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "待手术时间范围(&5)"
            Height          =   180
            Index           =   1
            Left            =   870
            TabIndex        =   0
            Top             =   855
            Width           =   1530
         End
         Begin VB.Label lbl 
            Caption         =   "在手术工作站中的待手术以及已完成手术的时间范围分别按如下设置进行搜索。"
            Height          =   405
            Index           =   11
            Left            =   780
            TabIndex        =   8
            Top             =   360
            Width           =   3840
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   165
            Picture         =   "frmOpsStationPara.frx":0028
            Top             =   240
            Width           =   480
         End
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "自动刷新(&1)         分"
         Height          =   180
         Index           =   8
         Left            =   105
         TabIndex        =   25
         Top             =   4725
         Width           =   1980
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3480
      TabIndex        =   4
      Top             =   5265
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4650
      TabIndex        =   5
      Top             =   5265
      Width           =   1100
   End
End
Attribute VB_Name = "frmOpsStationPara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mblnOK As Boolean
Private mlngLoop As Long
Private mfrmMain As Object
Private mstrPrivs As String

Public Function ShowPara(ByVal frmMain As Object, ByVal strPrivs As String) As Boolean
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim strPar As String
    
    Dim objCbo As Object
    
    Dim intLoop As Integer
    
    mblnOK = False
    mstrPrivs = strPrivs
    
    Set mfrmMain = frmMain
    '初始化
    '------------------------------------------------------------------------------------------------------------------
    For mlngLoop = 0 To 1
        cbo(mlngLoop).AddItem "今  天"
        cbo(mlngLoop).AddItem "昨  天"
        cbo(mlngLoop).AddItem "明  天"
        cbo(mlngLoop).AddItem "后  天"
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
    Next
    
    '缺省药房
    '------------------------------------------------------------------------------------------------------------------
    cbo(2).AddItem "手工选择"
    cbo(3).AddItem "手工选择"
    cbo(4).AddItem "手工选择"
    cbo(5).AddItem "手工选择"
    cbo(6).AddItem "手工选择"
    cbo(7).AddItem "手工选择"
    cbo(8).AddItem "手工选择"
    cbo(9).AddItem "手工选择"
    strSQL = _
        "Select Distinct A.ID,A.编码,A.名称,B.工作性质,B.服务对象" & _
        " From 部门表 A,部门性质说明 B " & _
        " Where (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
        " And B.部门ID=A.ID And B.服务对象 IN(1,2,3)" & _
        " And B.工作性质 in('西药房','成药房','中药房','发料部门')" & _
        " Order by A.编码"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    For intLoop = 1 To rsTmp.RecordCount
        If rsTmp!工作性质 = "西药房" Then
            Set objCbo = IIf(rsTmp!服务对象 = 1, cbo(2), IIf(rsTmp!服务对象 = 2, cbo(5), Nothing))
        End If
        If rsTmp!工作性质 = "成药房" Then
            Set objCbo = IIf(rsTmp!服务对象 = 1, cbo(3), IIf(rsTmp!服务对象 = 2, cbo(6), Nothing))
        End If
        If rsTmp!工作性质 = "中药房" Then
            Set objCbo = IIf(rsTmp!服务对象 = 1, cbo(4), IIf(rsTmp!服务对象 = 2, cbo(7), Nothing))
        End If
        If objCbo Is Nothing Then
            If rsTmp!工作性质 = "西药房" Then
                cbo(2).AddItem rsTmp!名称
                cbo(2).ItemData(cbo(2).NewIndex) = rsTmp!ID
                cbo(5).AddItem rsTmp!名称
                cbo(5).ItemData(cbo(5).NewIndex) = rsTmp!ID
            ElseIf rsTmp!工作性质 = "成药房" Then
                cbo(3).AddItem rsTmp!名称
                cbo(3).ItemData(cbo(3).NewIndex) = rsTmp!ID
                cbo(6).AddItem rsTmp!名称
                cbo(6).ItemData(cbo(6).NewIndex) = rsTmp!ID
            ElseIf rsTmp!工作性质 = "中药房" Then
                cbo(4).AddItem rsTmp!名称
                cbo(4).ItemData(cbo(4).NewIndex) = rsTmp!ID
                cbo(7).AddItem rsTmp!名称
                cbo(7).ItemData(cbo(7).NewIndex) = rsTmp!ID
            ElseIf rsTmp!工作性质 = "发料部门" Then
                cbo(8).AddItem rsTmp!名称
                cbo(8).ItemData(cbo(8).NewIndex) = rsTmp!ID
            End If
        Else
            objCbo.AddItem rsTmp!名称
            objCbo.ItemData(objCbo.NewIndex) = rsTmp!ID
        End If
        rsTmp.MoveNext
    Next
    
    strSQL = _
        "Select Distinct A.ID,A.编码,A.名称,B.工作性质,B.服务对象" & _
        " From 部门表 A,部门性质说明 B " & _
        " Where (A.撤档时间=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.撤档时间 is NULL)" & _
        " And B.部门ID=A.ID And B.服务对象 IN(1,2,3)" & _
        " And B.工作性质 in('手术')" & _
        " Order by A.编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If rsTmp.BOF = False Then
        For intLoop = 1 To rsTmp.RecordCount
            cbo(9).AddItem rsTmp!名称
            cbo(9).ItemData(cbo(9).NewIndex) = rsTmp!ID
            rsTmp.MoveNext
        Next
    End If
    
    
    strPar = GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "门诊缺省西药房", "0")
    For intLoop = 0 To cbo(2).ListCount - 1
        If cbo(2).ItemData(intLoop) = Val(strPar) Then cbo(2).ListIndex = intLoop: Exit For
    Next
    
    strPar = GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "门诊缺省成药房", "0")
    For intLoop = 0 To cbo(3).ListCount - 1
        If cbo(3).ItemData(intLoop) = Val(strPar) Then cbo(3).ListIndex = intLoop: Exit For
    Next
    
    strPar = GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "门诊缺省中药房", "0")
    For intLoop = 0 To cbo(4).ListCount - 1
        If cbo(4).ItemData(intLoop) = Val(strPar) Then cbo(4).ListIndex = intLoop: Exit For
    Next
    
    strPar = GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "住院缺省西药房", "0")
    For intLoop = 0 To cbo(5).ListCount - 1
        If cbo(5).ItemData(intLoop) = Val(strPar) Then cbo(5).ListIndex = intLoop: Exit For
    Next
    
    strPar = GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "住院缺省成药房", "0")
    For intLoop = 0 To cbo(6).ListCount - 1
        If cbo(6).ItemData(intLoop) = Val(strPar) Then cbo(6).ListIndex = intLoop: Exit For
    Next
    
    strPar = GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "住院缺省中药房", "0")
    For intLoop = 0 To cbo(7).ListCount - 1
        If cbo(7).ItemData(intLoop) = Val(strPar) Then cbo(7).ListIndex = intLoop: Exit For
    Next
        
    strPar = GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "材料缺省部门", "0")
    For intLoop = 0 To cbo(8).ListCount - 1
        If cbo(8).ItemData(intLoop) = Val(strPar) Then cbo(8).ListIndex = intLoop: Exit For
    Next
    
    strPar = GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "治疗缺省部门", "0")
    For intLoop = 0 To cbo(9).ListCount - 1
        If cbo(9).ItemData(intLoop) = Val(strPar) Then cbo(9).ListIndex = intLoop: Exit For
    Next
    
    On Error Resume Next
    
    
    cbo(0).Text = GetPara("待手术时间范围", mfrmMain.模块号, , "今  天")
    cbo(1).Text = GetPara("已完手术时间范围", mfrmMain.模块号, , "今  天")
    
    txt.Text = GetPara("自动刷新间隔", mfrmMain.模块号, , "0")
    chk.Value = GetPara("发送时不自生成费用", mfrmMain.模块号, , "0")

    If cbo(0).ListIndex = -1 Then cbo(0).ListIndex = 0
    If cbo(1).ListIndex = -1 Then cbo(1).ListIndex = 0
    
    cbo(0).ForeColor = COLOR.公共模块色
    cbo(1).ForeColor = COLOR.公共模块色
    txt.ForeColor = COLOR.公共模块色
    lbl(11).ForeColor = COLOR.公共模块色
    lbl(1).ForeColor = COLOR.公共模块色
    lbl(3).ForeColor = COLOR.公共模块色
    lbl(8).ForeColor = COLOR.公共模块色
    fra.ForeColor = COLOR.公共模块色
    
    fra.Enabled = IsPrivs(mstrPrivs, "参数设置")
    cbo(0).Enabled = fra.Enabled
    cbo(1).Enabled = fra.Enabled
    udn.Enabled = fra.Enabled
    txt.Enabled = fra.Enabled
    
    Me.Show 1, frmMain
    
    ShowPara = mblnOK
    
End Function

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDeviceSetup_Click()
    Call zlCommFun.DeviceSetup(Me, ParamInfo.系统号, ParamInfo.模块号)
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name, Int((ParamInfo.系统号) / 100)
End Sub

Private Sub cmdOK_Click()
    Dim strPar As String, i As Long
    
    Call SetPara("待手术时间范围", cbo(0).Text, mfrmMain.模块号)
    Call SetPara("已完手术时间范围", cbo(1).Text, mfrmMain.模块号)
    
    Call SetPara("自动刷新间隔", Val(txt.Text), mfrmMain.模块号)
    
    Call SetPara("发送时不自生成费用", chk.Value, mfrmMain.模块号)
        
    If cbo(2).ListIndex = -1 Then
        MsgBox "请指定缺省的门诊西药房。", vbInformation, gstrSysName
        cbo(2).SetFocus: Exit Sub
    End If
    If cbo(3).ListIndex = -1 Then
        MsgBox "请指定缺省的门诊成药房。", vbInformation, gstrSysName
        cbo(3).SetFocus: Exit Sub
    End If
    If cbo(4).ListIndex = -1 Then
        MsgBox "请指定缺省的门诊中药房。", vbInformation, gstrSysName
        cbo(4).SetFocus: Exit Sub
    End If
    If cbo(5).ListIndex = -1 Then
        MsgBox "请指定缺省的住院西药房。", vbInformation, gstrSysName
        cbo(5).SetFocus: Exit Sub
    End If
    If cbo(6).ListIndex = -1 Then
        MsgBox "请指定缺省的住院成药房。", vbInformation, gstrSysName
        cbo(6).SetFocus: Exit Sub
    End If
    If cbo(7).ListIndex = -1 Then
        MsgBox "请指定缺省的住院中药房。", vbInformation, gstrSysName
        cbo(7).SetFocus: Exit Sub
    End If
    
    If cbo(8).ListIndex = -1 Then
        MsgBox "请指定缺省的材料执行部门。", vbInformation, gstrSysName
        cbo(8).SetFocus: Exit Sub
    End If
    If cbo(9).ListIndex = -1 Then
        MsgBox "请指定缺省的治疗执行部门。", vbInformation, gstrSysName
        cbo(9).SetFocus: Exit Sub
    End If
    
    '缺省药房
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName, "门诊缺省西药房", cbo(2).ItemData(cbo(2).ListIndex)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName, "门诊缺省成药房", cbo(3).ItemData(cbo(3).ListIndex)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName, "门诊缺省中药房", cbo(4).ItemData(cbo(4).ListIndex)
    
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName, "住院缺省西药房", cbo(5).ItemData(cbo(5).ListIndex)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName, "住院缺省成药房", cbo(6).ItemData(cbo(6).ListIndex)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName, "住院缺省中药房", cbo(7).ItemData(cbo(7).ListIndex)
    
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName, "材料缺省部门", cbo(8).ItemData(cbo(8).ListIndex)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName, "治疗缺省部门", cbo(9).ItemData(cbo(9).ListIndex)
    
    mblnOK = True

    
    Unload Me
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub tbs_Click(PreviousTab As Integer)
    tbs.ZOrder 0
End Sub

Private Sub txt_Validate(Cancel As Boolean)
    Cancel = Not StrIsValid(txt.Text, txt.MaxLength)
    If Cancel Then Exit Sub

    If Val(txt.Text) < udn.Min Or Val(txt.Text) > udn.Max Then
        Cancel = True
    End If
End Sub
