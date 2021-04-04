VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm个人帐户调整 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "个人帐户调整"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6690
   Icon            =   "frm个人帐户调整.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   6690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fra基本 
      Caption         =   "条件"
      Height          =   2415
      Left            =   150
      TabIndex        =   0
      Top             =   90
      Width           =   6405
      Begin VB.TextBox txt身份证号 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4170
         MaxLength       =   18
         TabIndex        =   17
         Top             =   1125
         Width           =   2085
      End
      Begin VB.ComboBox Cbo人员类别 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4170
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   720
         Width           =   2085
      End
      Begin VB.ComboBox cbo性别 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1515
         Width           =   1635
      End
      Begin VB.ComboBox cbo中心1 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4170
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   330
         Width           =   2085
      End
      Begin VB.TextBox txt姓名 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1230
         MaxLength       =   20
         TabIndex        =   7
         Top             =   1125
         Width           =   1635
      End
      Begin VB.TextBox txt住院次数 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4170
         MaxLength       =   2
         TabIndex        =   19
         Top             =   1515
         Width           =   855
      End
      Begin VB.TextBox txt医保号 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1230
         MaxLength       =   20
         TabIndex        =   5
         Top             =   720
         Width           =   1635
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "…"
         Height          =   240
         Left            =   2580
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   360
         Width           =   255
      End
      Begin VB.TextBox txt帐户余额 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4170
         MaxLength       =   16
         TabIndex        =   21
         Top             =   1905
         Width           =   1545
      End
      Begin MSComCtl2.DTPicker dtp出生日期1 
         Height          =   300
         Left            =   1230
         TabIndex        =   11
         Top             =   1905
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   529
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy年MM月dd日"
         Format          =   86245379
         CurrentDate     =   36526
      End
      Begin VB.TextBox txt卡号 
         Height          =   300
         Left            =   1230
         MaxLength       =   20
         TabIndex        =   2
         Top             =   330
         Width           =   1635
      End
      Begin VB.Label lbl身份证号 
         AutoSize        =   -1  'True
         Caption         =   "身份证号(&I)"
         Enabled         =   0   'False
         Height          =   180
         Left            =   3090
         TabIndex        =   16
         Top             =   1185
         Width           =   990
      End
      Begin VB.Label lbl单位 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "元"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   5760
         TabIndex        =   22
         Top             =   1965
         Width           =   180
      End
      Begin VB.Label lbl人员类别1 
         AutoSize        =   -1  'True
         Caption         =   "人员类别(&K)"
         Enabled         =   0   'False
         Height          =   180
         Left            =   3090
         TabIndex        =   14
         Top             =   780
         Width           =   990
      End
      Begin VB.Label lbl出生日期1 
         AutoSize        =   -1  'True
         Caption         =   "出生日期(&B)"
         Enabled         =   0   'False
         Height          =   180
         Left            =   150
         TabIndex        =   10
         Top             =   1965
         Width           =   990
      End
      Begin VB.Label lbl性别 
         AutoSize        =   -1  'True
         Caption         =   "性别(&X)"
         Enabled         =   0   'False
         Height          =   180
         Left            =   510
         TabIndex        =   8
         Top             =   1575
         Width           =   630
      End
      Begin VB.Label lbl医保中心1 
         AutoSize        =   -1  'True
         Caption         =   "医保中心(&R)"
         Enabled         =   0   'False
         Height          =   180
         Left            =   3090
         TabIndex        =   12
         Top             =   390
         Width           =   990
      End
      Begin VB.Label lbl姓名 
         AutoSize        =   -1  'True
         Caption         =   "姓名(&N)"
         Enabled         =   0   'False
         Height          =   180
         Left            =   510
         TabIndex        =   6
         Top             =   1185
         Width           =   630
      End
      Begin VB.Label lbl住院次数 
         AutoSize        =   -1  'True
         Caption         =   "住院次数(&S)"
         Enabled         =   0   'False
         Height          =   180
         Left            =   3090
         TabIndex        =   18
         Top             =   1575
         Width           =   990
      End
      Begin VB.Label lbl医保号 
         AutoSize        =   -1  'True
         Caption         =   "医保号(&Y)"
         Enabled         =   0   'False
         Height          =   180
         Left            =   330
         TabIndex        =   4
         Top             =   780
         Width           =   810
      End
      Begin VB.Label lbl卡号 
         AutoSize        =   -1  'True
         Caption         =   "卡号(&D)"
         Height          =   180
         Left            =   510
         TabIndex        =   1
         Top             =   390
         Width           =   630
      End
      Begin VB.Label lbl帐户余额 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "帐户余额(&L)"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3090
         TabIndex        =   20
         Top             =   1965
         Width           =   990
      End
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   -270
      TabIndex        =   46
      Top             =   3780
      Width           =   7275
   End
   Begin VB.TextBox txt经办人 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1110
      TabIndex        =   39
      Top             =   3360
      Width           =   1275
   End
   Begin MSComctlLib.ImageList ImgLvw 
      Left            =   3060
      Top             =   1680
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
            Picture         =   "frm个人帐户调整.frx":06EA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fra规则 
      Caption         =   "规则"
      Height          =   705
      Left            =   150
      TabIndex        =   31
      Top             =   2550
      Width           =   6405
      Begin VB.TextBox txt浮动 
         Height          =   300
         Left            =   4170
         MaxLength       =   9
         TabIndex        =   36
         Top             =   270
         Width           =   1545
      End
      Begin VB.TextBox txt调整额 
         Height          =   300
         Left            =   1260
         MaxLength       =   15
         TabIndex        =   33
         Top             =   270
         Width           =   1605
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   5760
         TabIndex        =   37
         Top             =   330
         Width           =   90
      End
      Begin VB.Label lbl浮动 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "浮动(&F)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3450
         TabIndex        =   35
         Top             =   330
         Width           =   630
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "元"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   2910
         TabIndex        =   34
         Top             =   330
         Width           =   180
      End
      Begin VB.Label lbl调整额 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "调整额(&A)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   360
         TabIndex        =   32
         Top             =   330
         Width           =   810
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   240
      TabIndex        =   45
      Top             =   3990
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4140
      TabIndex        =   43
      Top             =   3990
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5370
      TabIndex        =   44
      Top             =   3990
      Width           =   1095
   End
   Begin VB.Frame fra批量 
      Caption         =   "条件"
      Height          =   2415
      Left            =   150
      TabIndex        =   23
      Top             =   90
      Width           =   6405
      Begin MSComctlLib.ListView lvw人员类别 
         Height          =   1125
         Left            =   1260
         TabIndex        =   30
         Top             =   1140
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   1984
         View            =   2
         Arrange         =   1
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImgLvw"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.ComboBox cbo中心2 
         Height          =   300
         Left            =   1260
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   330
         Width           =   3375
      End
      Begin MSComCtl2.DTPicker dtp出生日期2 
         Height          =   300
         Left            =   1260
         TabIndex        =   27
         Top             =   720
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   529
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "yyyy年MM月dd日"
         Format          =   86245379
         CurrentDate     =   36526
      End
      Begin VB.Label lbl人员类别2 
         AutoSize        =   -1  'True
         Caption         =   "人员类别(&K)"
         Height          =   180
         Left            =   180
         TabIndex        =   29
         Top             =   1170
         Width           =   990
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "这天以前出生的病人"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3000
         TabIndex        =   28
         Top             =   780
         Width           =   1620
      End
      Begin VB.Label lbl出生日期2 
         AutoSize        =   -1  'True
         Caption         =   "出生日期(&B)"
         Height          =   180
         Left            =   180
         TabIndex        =   26
         Top             =   780
         Width           =   990
      End
      Begin VB.Label lbl医保中心2 
         AutoSize        =   -1  'True
         Caption         =   "医保中心(&R)"
         Height          =   180
         Left            =   180
         TabIndex        =   24
         Top             =   390
         Width           =   990
      End
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "核对(&T)"
      Height          =   350
      Left            =   2910
      TabIndex        =   42
      Top             =   3990
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txt说明 
      Height          =   300
      Left            =   3540
      MaxLength       =   200
      TabIndex        =   41
      Top             =   3360
      Width           =   3015
   End
   Begin VB.Label lbl说明 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "说明(&M)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   2820
      TabIndex        =   40
      Top             =   3420
      Width           =   630
   End
   Begin VB.Label lbl经办人 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "经办人(&P)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   225
      TabIndex        =   38
      Top             =   3420
      Width           =   810
   End
End
Attribute VB_Name = "frm个人帐户调整"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mint险类 As Integer
Private mint操作模式 As Integer                     '1-单个调整;2-批量调整;3-修改;4-查阅
Private mlng记录ID As Long                          '指修改记录的ID
Private mblnOK As Boolean                           '是否更新数据库
Private mbln校验 As Boolean
Private mblnStart As Boolean

Public Function ShowME(ByVal frmParent As Object, ByVal int操作模式 As Integer, _
ByVal int险类 As Integer, Optional ByVal lng记录ID As Long = 0) As Boolean
    mblnOK = False
    
    mint操作模式 = int操作模式
    mint险类 = int险类
    mlng记录ID = lng记录ID
    Me.Show 1, frmParent
    ShowME = mblnOK
End Function

Private Function InitFace() As Boolean
    Dim rsTemp As New ADODB.Recordset
    '填充基本的数据
    
    gstrSQL = "Select 名称,编码 ID From 性别"
    Call OpenRecordset(rsTemp, Me.Caption)
    Call zlControl.CboAddData(Me.cbo性别, rsTemp, True)
    Me.cbo性别.ListIndex = 0
    
    gstrSQL = "Select 名称,序号 ID From 保险中心目录 Where 险类=" & mint险类
    Call OpenRecordset(rsTemp, Me.Caption)
    Call zlControl.CboAddData(Me.cbo中心1, rsTemp, True)
    cbo中心1.ListIndex = 0
    cbo中心2.Clear
    cbo中心2.AddItem "所有医保中心"
    cbo中心2.ItemData(cbo中心2.NewIndex) = 0
    Call zlControl.CboAddData(Me.cbo中心2, rsTemp, False)
    Me.cbo中心2.ListIndex = 0
    
    gstrSQL = "Select 名称,序号 ID From 保险人群 Where 险类=" & mint险类
    Call OpenRecordset(rsTemp, Me.Caption)
    Call zlControl.CboAddData(Me.Cbo人员类别, rsTemp, True)
    Cbo人员类别.ListIndex = 0
    With rsTemp
        .MoveFirst
        lvw人员类别.ListItems.Clear
        lvw人员类别.ListItems.Add , "K_0", "所有人员类别", , 1
        Do While Not .EOF
            lvw人员类别.ListItems.Add , "K_" & !ID, !名称, , 1
            .MoveNext
        Loop
    End With
    txt经办人 = gstrUserName
    If mint操作模式 < 3 Then
        InitFace = True
        Exit Function
    End If
    
    '如果是修改或查阅，则要读出原始数据（修改时，帐户余额是修改前的余额）
    gstrSQL = "Select B.ID,A.中心,A.卡号,A.医保号,C.病人ID,C.姓名,A.病人ID, " & _
             " C.性别,C.出生日期,A.在职 人员类别,Nvl(A.帐户余额,0) 帐户余额,C.身份证号,A.退休证号, " & _
             " ltrim(to_char(B.金额,'900090000.00')) 金额,B.经办人,Nvl(D.住院次数累计,0) 本院,Nvl(D.外院住院次数,0) 外院, " & _
             " To_char(B.时间,'yyyy-MM-dd hh24:mi:ss') 时间,说明  " & _
             " From 保险帐户 A,帐户变动记录 B,病人信息 C , " & _
             " (Select * From 帐户年度信息 Where 年度=to_char(Sysdate,'yyyy')) D " & _
             " Where A.险类=B.险类 And A.病人ID=B.病人ID And A.病人ID=C.病人ID  " & _
             " And A.险类=D.险类(+) And A.病人ID=D.病人ID(+) And A.险类=" & mint险类 & " And B.ID=" & mlng记录ID
    Call OpenRecordset(rsTemp, Me.Caption)
    
    If rsTemp.EOF Then
        MsgBox "没找到指定的帐户变动记录，可能已经被其它操作员删除！", vbInformation, gstrSysName
        Exit Function
    End If
    Call WriteCons(rsTemp)
    If mint操作模式 = 3 Then
        InitFace = True
        Exit Function
    End If
    
    Call DisableCons
    InitFace = True
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Function Calc金额(ByVal cur帐户余额 As Currency) As Currency
    '计算实际的调整额
    If Val(txt调整额.Text) <> 0 Then
        Calc金额 = Val(txt调整额.Text)
    Else
        If Val(txt浮动.Text) < 0 Then
            Calc金额 = Val(cur帐户余额) * Abs(txt浮动.Text) / 100 * -1
        Else
            Calc金额 = Val(cur帐户余额) * Abs(txt浮动.Text) / 100
        End If
    End If
End Function

Private Sub cmdOK_Click()
    Dim lngNextID As Long, cur金额 As Currency
    Dim rsAccount As New ADODB.Recordset
    If Not ValidData Then Exit Sub
    If mint操作模式 = 4 Then
        Unload Me
        Exit Sub
    End If
    
    On Error GoTo errHand
    
    '如果调整额为零且浮动也为零，则退出
    If Val(txt调整额) = 0 And Val(txt浮动) = 0 Then
        MsgBox "请输入调整额或浮动比例！", vbInformation, gstrSysName
        txt调整额.SetFocus
        Exit Sub
    End If
    
    gcnOracle.BeginTrans
    Select Case mint操作模式
    Case 1
        '调整单个医保病人的帐户余额
        cur金额 = Calc金额(Val(txt帐户余额.Text))
        lngNextID = zlDatabase.GetNextID("帐户变动记录")
        gstrSQL = "ZL_帐户变动记录_INSERT(" & _
                  lngNextID & "," & mint险类 & ",1," & Val(txt卡号.Tag) & "," & _
                  cur金额 & ",'" & txt经办人.Text & "','" & txt说明.Text & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        
        Call 检查帐户信息_米易(Val(txt卡号.Tag), True, False)
    Case 2
        '批量调整
        gstrSQL = " Select A.病人ID,Nvl(A.帐户余额,0) 帐户余额" & _
                  " From 保险帐户 A,病人信息 B" & _
                  " Where A.病人ID=B.病人ID And Nvl(A.灰度级,0)<>9 And A.险类=" & mint险类 & GetSQL
        Call OpenRecordset(rsAccount, "统计记录数，以便更新")
        
        Do While Not rsAccount.EOF
            cur金额 = Calc金额(rsAccount!帐户余额)
            lngNextID = zlDatabase.GetNextID("帐户变动记录")
            gstrSQL = "ZL_帐户变动记录_INSERT(" & _
                      lngNextID & "," & mint险类 & ",1," & rsAccount!病人ID & "," & _
                      cur金额 & ",'" & txt经办人.Text & "','" & txt说明.Text & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            
            Call 检查帐户信息_米易(rsAccount!病人ID, True, False)
            rsAccount.MoveNext
        Loop
    Case 3
        '修改
        cur金额 = Calc金额(Val(txt帐户余额.Text))
        gstrSQL = "ZL_帐户变动记录_UPDATE(" & _
            mlng记录ID & "," & cur金额 & ",'" & txt经办人.Text & "','" & txt说明.Text & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        
        Call 检查帐户信息_米易(Val(txt卡号.Tag), True, False)
    End Select
    
    mblnOK = True
    '只对调整的帐户进行打印，修改或删除的或帐户新增时初始的帐户余额，请用户自己在管理界面把过程清单打出来
    If mint操作模式 = 1 Then
        Call zl9Report.ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1607", Me, "记录ID=" & lngNextID, 1)
    End If
    
    gcnOracle.CommitTrans
    
    '如果是修改则退出
    If mint操作模式 = 3 Then
        Unload Me
        Exit Sub
    End If
    
    '为继续新增做准备工作
    Call ClearAllCons
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
    gcnOracle.RollbackTrans
End Sub

Private Sub cmdSelect_Click()
    gstrSQL = " Select A.病人ID as ID,A.卡号,A.医保号,B.姓名,B.性别,B.出生日期,B.身份证号,C.序号 as 中心ID " & _
            " ,A.人员身份,A.单位编码,A.病种ID,D.名称 as 病种,A.在职 as 在职ID,A.退休证号,A.帐户余额" & _
            " From 保险帐户 A,病人信息 B,保险中心目录 C,保险病种 D" & _
            "  where A.病人ID=B.病人ID and A.险类=" & mint险类 & _
            "  and A.险类=C.险类 and A.中心=C.序号 and A.病种ID=D.ID(+)"
    
    Call Get帐户情况
    zlControl.TxtSelAll txt卡号
    txt卡号.SetFocus
End Sub

Private Sub cmdTest_Click()
    Dim strMsg As String
    Dim rsAccount As New ADODB.Recordset
    '用于批量执行前，按当前设定的条件，统计共有多少医保病人会调整余额，以便操作员进行对比
    
    '统计共有多少医保病人的帐户余额会调整
    gstrSQL = " Select Count(*) 记录数" & _
              " From 保险帐户 A,病人信息 B" & _
              " Where A.病人ID=B.病人ID And Nvl(A.灰度级,0)<>9 And A.险类=" & mint险类 & GetSQL
    Call OpenRecordset(rsAccount, "统计记录数，以便核对条件是否正确")
    
    If rsAccount!记录数 = 0 Then
        strMsg = "没有符合条件的记录！"
        mbln校验 = False
        cmdOK.Enabled = False
        MsgBox strMsg, vbInformation, gstrSysName
        Exit Sub
    Else
        strMsg = "按当前设定的条件共统计出： " & rsAccount!记录数 & " 个医保帐户的余额将会调整"
    End If
    
    '如果是直接输入的调整额，统计是否会存在调整为负的情况
    If Val(txt调整额.Text) < 0 Then
        gstrSQL = " Select Count(*) 记录数" & _
                  " From 保险帐户 A,病人信息 B" & _
                  " Where A.病人ID=B.病人ID And Nvl(A.灰度级,0)<>9 And A.险类=" & mint险类 & GetSQL & _
                  " And Nvl(帐户余额,0)<" & Val(Abs(txt调整额))
        Call OpenRecordset(rsAccount, "统计出帐户余额会调整为负数的所有记录")
        If rsAccount!记录数 <> 0 Then
            strMsg = strMsg & vbCrLf & "但其中有" & rsAccount!记录数 & "个医保帐户的余额会调整为负！"
            mbln校验 = False
            cmdOK.Enabled = False
            MsgBox strMsg, vbInformation, gstrSysName
        End If
    End If
    
    cmdOK.Enabled = True
    MsgBox strMsg, vbInformation, gstrSysName
End Sub

Private Function ValidData() As Boolean
    '检测输入数据合法性
    If mint操作模式 <> 2 Then
        If Val(txt卡号.Tag) = 0 Then
            MsgBox "请输入医保病人的卡号！", vbInformation, gstrSysName
            txt卡号.SetFocus
            Exit Function
        End If
    End If
    
    If Val(txt浮动.Text) <> 0 Then
        If Abs(txt浮动.Text) > 100 Then
            MsgBox "浮动额不能大于100%！", vbInformation, gstrSysName
            txt浮动.SetFocus
            Exit Function
        End If
    End If
    If Val(txt调整额.Text) <> 0 Then
        If Abs(txt调整额.Text) > 100000000000000# Then
            MsgBox "调整额超过最大值！", vbInformation, gstrSysName
            txt调整额.SetFocus
            Exit Function
        End If
        
        '调整额不能大于帐户余额
        If mint操作模式 <> 2 Then
            If Val(txt帐户余额.Text) + Val(txt调整额.Text) < 0 Then
                MsgBox "调整额不能大于帐户余额！", vbInformation, gstrSysName
                txt调整额.SetFocus
                Exit Function
            End If
        End If
    End If
    
    If Trim(txt经办人.Text) = "" Then
        MsgBox "请先设置了当前用户对应的人员！", vbInformation, gstrSysName
        Exit Function
    End If
    If zlCommFun.ActualLen(txt说明.Text) > 200 Then
        MsgBox "说明的内容超长（最多100个汉字或200个字符）！", vbInformation, gstrSysName
        txt说明.SetFocus
        Exit Function
    End If
    
    ValidData = True
End Function

Private Function GetSQL() As String
    Dim strSQL As String, str人员类别 As String
    Dim intItem As Integer
    Dim bln中心 As Boolean
    Dim rs中心 As New ADODB.Recordset
    '返回用户设定的SQL串
    
    bln中心 = 存在中心(mint险类)
    strSQL = "": str人员类别 = ""
    If cbo中心2.ListIndex <> 0 Then
        If bln中心 Then
            strSQL = strSQL & " And A.中心=" & cbo中心2.ItemData(cbo中心2.ListIndex)
        End If
    End If
    If Not IsNull(dtp出生日期2.Value) Then
        strSQL = strSQL & " And B.出生日期<to_date('" & Format(dtp出生日期2.Value, "yyyy-MM-dd") & "','yyyy-MM-dd')"
    End If
    With lvw人员类别
        If Not .ListItems(1).Checked Then
            For intItem = 2 To .ListItems.Count
                If .ListItems(intItem).Checked Then str人员类别 = str人员类别 & IIf(str人员类别 = "", "", ",") & Mid(.ListItems(intItem).Key, 3)
            Next
        End If
    End With
    If str人员类别 <> "" Then
        strSQL = strSQL & " And A.在职 in (" & str人员类别 & ")"
    End If
    GetSQL = strSQL
End Function

Private Sub Form_Activate()
    If mblnStart = False Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    mblnStart = False
    fra批量.Visible = (mint操作模式 = 2)
    fra基本.Visible = Not (mint操作模式 = 2)
    cmdTest.Visible = (mint操作模式 = 2)
    If cmdTest.Visible Then cmdOK.Enabled = False
    
    mblnStart = InitFace
End Sub

Private Sub lvw人员类别_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim intItems As Integer
    Dim blnState As Boolean
    
    If Item.Key = "K_0" Then
        For intItems = 1 To lvw人员类别.ListItems.Count
            lvw人员类别.ListItems(intItems).Checked = Item.Checked
        Next
    Else
        '如果余下的全部选择或全部未选择，则更新第一项的状态
        blnState = lvw人员类别.ListItems(2).Checked             '至少有一种人员类别
        For intItems = 2 To lvw人员类别.ListItems.Count
            If blnState <> lvw人员类别.ListItems(intItems).Checked Then
                lvw人员类别.ListItems(1).Checked = False
                Exit Sub
            End If
        Next
        If blnState Then
            lvw人员类别.ListItems(1).Checked = blnState
        Else
            '必须选择一种人员类别
            lvw人员类别.ListItems(2).Checked = True
        End If
    End If
End Sub

Private Sub txt调整额_Change()
    On Error Resume Next
    If Me.ActiveControl.Name = "txt调整额" Then txt浮动.Text = ""
End Sub

Private Sub txt调整额_GotFocus()
    Call zlControl.TxtSelAll(txt调整额)
End Sub

Private Sub txt调整额_KeyPress(KeyAscii As Integer)
    If Not (InStr(1, "0123456789.-", Chr(KeyAscii)) <> 0 Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

Private Sub txt调整额_Validate(Cancel As Boolean)
    txt调整额.Text = Format(txt调整额.Text, "#####0.00;-#####0.00; ;")
End Sub

Private Sub txt浮动_Change()
    On Error Resume Next
    If Me.ActiveControl.Name = "txt浮动" Then txt调整额.Text = ""
End Sub

Private Sub txt浮动_GotFocus()
    Call zlControl.TxtSelAll(txt浮动)
End Sub

Private Sub txt浮动_KeyPress(KeyAscii As Integer)
    If Not (InStr(1, "0123456789.-", Chr(KeyAscii)) <> 0 Or KeyAscii = vbKeyBack) Then KeyAscii = 0
End Sub

Private Sub txt浮动_Validate(Cancel As Boolean)
    txt浮动.Text = Format(txt浮动.Text, "#####0.00;-#####0.00; ;")
End Sub

Private Sub ClearAllCons()
    Select Case mint操作模式
    Case 1
        txt卡号.Text = ""
        txt卡号.Tag = ""
        txt医保号.Text = ""
        txt姓名.Text = ""
        txt身份证号.Text = ""
        txt住院次数.Text = "0/0"
        txt帐户余额.Text = ""
    Case 2
        
    End Select
    
    txt调整额.Text = ""
    txt浮动.Text = ""
    txt说明.Text = ""
End Sub

Private Sub txt卡号_GotFocus()
    zlCommFun.OpenIme True
    zlControl.TxtSelAll txt卡号
End Sub

Private Sub txt卡号_KeyPress(KeyAscii As Integer)
    Dim strCode As String
    Dim str条件 As String
    Dim rsTemp As New ADODB.Recordset
    
    If Len(txt卡号.Text) = txt卡号.MaxLength Or KeyAscii = vbKeyReturn Then
        strCode = Replace(Trim(UCase(txt卡号.Text)), "'", "")
        If strCode = "" Then Exit Sub
        
        If IsNumeric(strCode) And IsNumeric(Left(strCode, 1)) Then '刷卡
            str条件 = " and A.卡号='" & strCode & "' and A.中心=" & cbo中心1.ItemData(cbo中心1.ListIndex)
        ElseIf (Left(strCode, 1) = "A" Or Left(strCode, 1) = "-") And IsNumeric(Mid(strCode, 2)) Then '病人ID
            str条件 = " and A.病人ID=" & Mid(strCode, 2)
        ElseIf (Left(strCode, 1) = "B" Or Left(strCode, 1) = "+") And IsNumeric(Mid(strCode, 2)) Then '住院号(对住(过)院的病人)
            str条件 = " and B.住院号=" & Mid(strCode, 2)
        ElseIf (Left(strCode, 1) = "D" Or Left(strCode, 1) = "*") And IsNumeric(Mid(strCode, 2)) Then '门诊号(仅对门诊病人)
            str条件 = " and B.门诊号=" & Mid(strCode, 2)
        Else '当作姓名
            str条件 = " and A.卡号='" & strCode & "'"
        End If
    
        gstrSQL = " Select A.病人ID as ID,A.卡号,A.医保号,B.姓名,B.性别,B.出生日期,B.身份证号,C.序号 as 中心ID " & _
                " ,A.人员身份,A.单位编码,A.病种ID,D.名称 as 病种,A.在职 as 在职ID,A.退休证号,A.帐户余额" & _
                " From 保险帐户 A,病人信息 B,保险中心目录 C,保险病种 D" & _
                "  where A.病人ID=B.病人ID and A.险类=" & mint险类 & _
                "  and A.险类=C.险类 and A.中心=C.序号 And Nvl(A.灰度级,0)<>9 and A.病种ID=D.ID(+)" & str条件
        
        Call Get帐户情况
    End If
End Sub

Private Sub txt卡号_LostFocus()
    zlCommFun.OpenIme False
End Sub

Private Sub Get帐户情况()
'从已经存在的记录中读出帐户信息
    Dim rs帐户 As ADODB.Recordset
    Dim rsTemp As New ADODB.Recordset
    Dim lngIndex As Long
    
    Set rs帐户 = frmPubSel.ShowSelect(Me, gstrSQL, 0, "保险帐户", , txt卡号.Text, "", False, True)
    If Not rs帐户 Is Nothing Then
        txt卡号.Text = rs帐户("卡号")
        txt卡号.Tag = rs帐户!ID
        '其它可用的数据
        txt医保号.Text = IIf(IsNull(rs帐户("医保号")), "", rs帐户("医保号"))
        txt姓名.Text = IIf(IsNull(rs帐户("姓名")), "", rs帐户("姓名"))
        txt身份证号.Text = IIf(IsNull(rs帐户("身份证号")), "", rs帐户("身份证号"))
        
        Call SetComboByText(cbo性别, IIf(IsNull(rs帐户("性别")), "", rs帐户("性别")), True)
        If IsNull(rs帐户("出生日期")) = False Then
            dtp出生日期1.Value = rs帐户("出生日期")
        End If
        
        For lngIndex = 0 To cbo中心1.ListCount - 1
            If cbo中心1.ItemData(lngIndex) = rs帐户("中心ID") Then
                cbo中心1.ListIndex = lngIndex
                Exit For
            End If
        Next
        txt帐户余额 = Format(rs帐户!帐户余额, "#####0.00;-#####0.00; ;")
        txt帐户余额.Enabled = False
        
        '再读出帐户年度信息
        gstrSQL = "select * from 帐户年度信息 where 险类=" & mint险类 & _
            " and 病人ID=" & rs帐户("ID") & " and 年度=" & Format(zlDatabase.Currentdate, "yyyy")
        Call OpenRecordset(rsTemp, Me.Caption)
        
        If rsTemp.EOF = False Then
            '设置帐户情况
            txt住院次数.Text = Nvl(rsTemp("住院次数累计"), "0") & "/" & Nvl(rsTemp("外院住院次数"), "0")
        Else
            txt住院次数.Text = "0/0"
        End If
    End If
End Sub

Private Sub WriteCons(ByVal rsObj As ADODB.Recordset)
    Dim cur帐户余额 As Currency, cur调整额 As Currency
    
    '将数据写入界面
    txt卡号.Text = rsObj!卡号
    txt卡号.Tag = rsObj!病人ID
    txt医保号 = rsObj!医保号
    txt姓名 = rsObj!姓名
    Call zlControl.CboLocate(cbo性别, rsObj!性别)
    dtp出生日期1.Value = Format(rsObj!出生日期, "yyyy-MM-dd")
    Call zlControl.CboLocate(cbo中心1, rsObj!中心, True)
    Call zlControl.CboLocate(Cbo人员类别, rsObj!人员类别, True)
    txt身份证号.Text = Nvl(rsObj!身份证号)
    txt住院次数.Text = Nvl(rsObj!本院, "0") & "/" & Nvl(rsObj!外院, 0)
    
    cur帐户余额 = Nvl(rsObj!帐户余额, 0)
    cur调整额 = Nvl(rsObj!金额, 0)
    cur帐户余额 = cur帐户余额 - cur调整额
    txt帐户余额.Text = Format(cur帐户余额, "#####0.00;-#####0.00; ;")
    txt调整额.Text = Format(cur调整额, "#####0.00;-#####0.00; ;")
    txt说明.Text = Nvl(rsObj!说明)
End Sub

Private Sub DisableCons()
    txt卡号.Enabled = False
    cmdSelect.Enabled = False
    txt调整额.Enabled = False
    txt浮动.Enabled = False
    txt说明.Enabled = False
    cmdTest.Visible = False
    cmdOK.Visible = False
    cmdCancel.Caption = "确定(&O)"
End Sub

Private Sub txt说明_GotFocus()
    Call zlControl.TxtSelAll(txt说明)
End Sub
