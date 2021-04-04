VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSetExpence 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6345
   Icon            =   "frmSetExpence.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3585
      TabIndex        =   1
      Top             =   4830
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4815
      TabIndex        =   2
      Top             =   4830
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   360
      TabIndex        =   3
      Top             =   4830
      Width           =   1100
   End
   Begin TabDlg.SSTab stab 
      Height          =   4650
      Left            =   150
      TabIndex        =   0
      Top             =   90
      Width           =   6030
      _ExtentX        =   10636
      _ExtentY        =   8202
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      WordWrap        =   0   'False
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "结算参数"
      TabPicture(0)   =   "frmSetExpence.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdBill"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraTitle"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "opt(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "opt(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cbo"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "chk"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      Begin VB.CheckBox chk 
         Caption         =   "对零费用进行结帐"
         Height          =   210
         Left            =   225
         TabIndex        =   11
         Top             =   4065
         Width           =   2805
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Left            =   1455
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   3585
         Width           =   1695
      End
      Begin VB.OptionButton opt 
         Caption         =   "门诊收费票据"
         Height          =   210
         Index           =   1
         Left            =   1845
         TabIndex        =   8
         Top             =   555
         Width           =   1530
      End
      Begin VB.OptionButton opt 
         Caption         =   "住院结帐票据"
         Height          =   210
         Index           =   0
         Left            =   255
         TabIndex        =   7
         Top             =   570
         Value           =   -1  'True
         Width           =   1530
      End
      Begin VB.Frame fraTitle 
         Caption         =   "本地共用结算票据"
         Height          =   2535
         Left            =   195
         TabIndex        =   5
         Top             =   975
         Width           =   5640
         Begin MSComctlLib.ListView lvwBill 
            Height          =   2220
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   5385
            _ExtentX        =   9499
            _ExtentY        =   3916
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   0   'False
            HideSelection   =   0   'False
            Checkboxes      =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            SmallIcons      =   "img16"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   1
            NumItems        =   4
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "领用人"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   1
               Text            =   "领用日期"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "号码范围"
               Object.Width           =   2910
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   3
               Text            =   "剩余"
               Object.Width           =   1235
            EndProperty
         End
         Begin MSComctlLib.ImageList img16 
            Left            =   735
            Top             =   510
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   1
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   1
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmSetExpence.frx":0028
                  Key             =   ""
               EndProperty
            EndProperty
         End
      End
      Begin VB.CommandButton cmdBill 
         Caption         =   "结算票据设置(&P)"
         Height          =   350
         Left            =   4095
         TabIndex        =   4
         Top             =   4125
         Width           =   1740
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "默认结算方式"
         Height          =   180
         Left            =   165
         TabIndex        =   10
         Top             =   3660
         Width           =   1080
      End
   End
End
Attribute VB_Name = "frmSetExpence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit '要求变量声明

Private mblnOK As Boolean
Private mint票种 As Integer

Public Function ShowParameter(ByVal frmMain As Object) As Boolean
    
    mblnOK = False
    
    Me.Show 1, frmMain
    
    ShowParameter = mblnOK
    
End Function

Private Sub cmdBill_Click()
    Call ReportPrintSet(gcnOracle, glngSys, "ZL1_BILL_1862", Me)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    '
End Sub

Private Sub cmdOK_Click()
    Dim lngLoop As Long
    
    '保存参数注册信息
    
    On Error Resume Next
    
    '本地共用结帐票据
    If opt(0).Value Then
        SaveSetting "ZLSOFT", "公共模块\" & App.ProductName, "共用结帐票据类型", 3
    Else
        SaveSetting "ZLSOFT", "公共模块\" & App.ProductName, "共用结帐票据类型", 1
    End If
    
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName, "对零费用进行结帐", chk.Value
    Call SaveSetting("ZLSOFT", "公共模块\" & App.ProductName, "缺省结算方式", cbo.Text)
    
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName, "共用结帐票据批次", 0
    For lngLoop = 1 To lvwBill.ListItems.Count
        If lvwBill.ListItems(lngLoop).Checked Then
            SaveSetting "ZLSOFT", "公共模块\" & App.ProductName, "共用结帐票据批次", Mid(lvwBill.ListItems(lngLoop).Key, 2)
        End If
    Next
            
    mblnOK = True
    Unload Me
End Sub

Private Function ReadBills(ByVal int票种 As Integer) As Boolean

    Dim rsTmp As New ADODB.Recordset
    Dim objItem As ListItem, blnBill As Boolean
    Dim lngLoop As Long
    Dim strTmp As String
    
    lvwBill.ListItems.Clear
    '读取可用公用结帐领用
    gstrSQL = "Select * From 票据领用记录 Where 票种=[1] And 使用方式=2 And 剩余数量>0 Order by 剩余数量 Desc"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, int票种)
    blnBill = False
    If Not rsTmp.EOF Then
        For lngLoop = 1 To rsTmp.RecordCount
            Set objItem = lvwBill.ListItems.Add(, "_" & rsTmp!ID, rsTmp!领用人, , 1)
            objItem.SubItems(1) = Format(rsTmp!登记时间, "yyyy-MM-dd")
            objItem.SubItems(2) = rsTmp!开始号码 & "," & rsTmp!终止号码
            objItem.SubItems(3) = rsTmp!剩余数量
            If rsTmp!ID = GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "共用结帐票据批次", 0) Then
                objItem.Checked = True
                objItem.Selected = True
                blnBill = True
            End If
            rsTmp.MoveNext
        Next
    End If
    If Not blnBill Then SaveSetting "ZLSOFT", "公共模块\" & App.ProductName, "共用结帐票据批次", 0
    
    
    strTmp = Trim(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "缺省结算方式", ""))

    gstrSQL = "SELECT A.结算方式 " & _
                "from 结算方式应用 A,结算方式 B where A.结算方式=B.名称 AND A.应用场合=[1] AND 性质 in (1,2)"
                    
    If int票种 = 1 Then
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, "收费")
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, "结帐")
    End If
    
    cbo.Clear
    
    If Not rsTmp.EOF Then
        Do While Not rsTmp.EOF
            cbo.AddItem zlCommFun.NVL(rsTmp("结算方式").Value)
            If strTmp = zlCommFun.NVL(rsTmp("结算方式").Value) Then
                cbo.ListIndex = cbo.NewIndex
            End If
            
            rsTmp.MoveNext
        Loop
    End If
    
End Function

Private Sub Form_Load()
    
    chk.Value = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "对零费用进行结帐", 1))
    mint票种 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "共用结帐票据类型", 3))
    If mint票种 = 1 Then
        opt(1).Value = True
    Else
        opt(0).Value = True
    End If
    
    Call ReadBills(mint票种)
    
End Sub

Private Sub lvwBill_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim lngLoop As Long
    
    For lngLoop = 1 To lvwBill.ListItems.Count
        If lvwBill.ListItems(lngLoop).Key <> Item.Key Then lvwBill.ListItems(lngLoop).Checked = False
    Next
    
    Item.Selected = True
    
End Sub

Private Sub opt_Click(Index As Integer)
    Call ReadBills(IIf(Index = 0, 3, 1))
End Sub
