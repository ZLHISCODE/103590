VERSION 5.00
Begin VB.Form frmSet沈阳 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "运行参数设置"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4845
   Icon            =   "frmSet沈阳.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   4845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.ComboBox cbo适用地区 
      Height          =   300
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   4140
      Width           =   1785
   End
   Begin VB.CheckBox chk是否允许在院病人进行门诊业务 
      Caption         =   "是否允许在院病人进行门诊业务"
      Height          =   255
      Left            =   360
      TabIndex        =   17
      Top             =   3810
      Width           =   2895
   End
   Begin VB.CheckBox chk挂号 
      Alignment       =   1  'Right Justify
      Caption         =   "挂号(&R)"
      Height          =   195
      Left            =   300
      TabIndex        =   13
      Top             =   2940
      Width           =   945
   End
   Begin VB.Frame fra挂号 
      Enabled         =   0   'False
      Height          =   705
      Left            =   210
      TabIndex        =   14
      Top             =   2970
      Width           =   3165
      Begin VB.ComboBox cbo收入项目 
         Height          =   300
         Left            =   1380
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   240
         Width           =   1515
      End
      Begin VB.Label lbl个人帐户支付 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "个人帐户支付"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   210
         TabIndex        =   15
         Top             =   300
         Width           =   1080
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3540
      TabIndex        =   21
      Top             =   840
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3540
      TabIndex        =   20
      Top             =   360
      Width           =   1100
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   3
      Left            =   2220
      MaxLength       =   2
      TabIndex        =   12
      Top             =   2640
      Width           =   855
   End
   Begin VB.Frame fra医保服务器 
      Caption         =   "医保服务器"
      Height          =   2385
      Left            =   210
      TabIndex        =   0
      Top             =   120
      Width           =   3165
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   5
         Left            =   1260
         MaxLength       =   40
         PasswordChar    =   "*"
         TabIndex        =   10
         Top             =   1890
         Width           =   1635
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   4
         Left            =   1260
         MaxLength       =   40
         TabIndex        =   8
         Top             =   1500
         Width           =   1635
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   0
         Left            =   1260
         MaxLength       =   40
         TabIndex        =   2
         Top             =   330
         Width           =   1635
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   1260
         MaxLength       =   40
         TabIndex        =   4
         Top             =   720
         Width           =   1635
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   1260
         MaxLength       =   40
         TabIndex        =   6
         Top             =   1110
         Width           =   1635
      End
      Begin VB.Label lblEdit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "密码(&W)"
         Height          =   180
         Index           =   5
         Left            =   570
         TabIndex        =   9
         Top             =   1950
         Width           =   630
      End
      Begin VB.Label lblEdit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "操作员(&U)"
         Height          =   180
         Index           =   4
         Left            =   390
         TabIndex        =   7
         Top             =   1560
         Width           =   810
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "地址(&A)"
         Height          =   180
         Index           =   0
         Left            =   570
         TabIndex        =   1
         Top             =   390
         Width           =   630
      End
      Begin VB.Label lblEdit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "端口号(&P)"
         Height          =   180
         Index           =   1
         Left            =   390
         TabIndex        =   3
         Top             =   780
         Width           =   810
      End
      Begin VB.Label lblEdit 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "入口程序(&S)"
         Height          =   180
         Index           =   2
         Left            =   210
         TabIndex        =   5
         Top             =   1170
         Width           =   990
      End
   End
   Begin VB.Label lbl适用地区 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "适用地区(&K)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   360
      TabIndex        =   18
      Top             =   4200
      Width           =   990
   End
   Begin VB.Label lblEdit 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "IC卡读写设备端口号(&I)"
      Height          =   180
      Index           =   3
      Left            =   270
      TabIndex        =   11
      Top             =   2700
      Width           =   1890
   End
End
Attribute VB_Name = "frmSet沈阳"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Modified By 朱玉宝 地区：长沙 原因：这个程序加上了登录操作员及登录口令
Private Enum 参数
    地址 = 0
    端口号
    入口程序
    IC设备端口
    登录操作员
    登录口令
End Enum
Private mblnReturn As Boolean

Private Sub chk挂号_Click()
    On Error Resume Next
    fra挂号.Enabled = (chk挂号.Value = 1)
    If fra挂号.Enabled Then
        cbo收入项目.SetFocus
    Else
        cbo收入项目.ListIndex = 0
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHand
    
    If Not Valid Then Exit Sub
    gcnOracle.BeginTrans
    
    '删除已经数据
    gstrSQL = "zl_保险参数_Delete(" & TYPE_沈阳市 & ",null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '新增参数数据
    gstrSQL = "zl_保险参数_Insert(" & TYPE_沈阳市 & ",NULL,'服务器地址','" & txtEdit(参数.地址).Text & "',1)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & TYPE_沈阳市 & ",NULL,'服务器端口号','" & txtEdit(参数.端口号).Text & "',2)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & TYPE_沈阳市 & ",NULL,'服务器入口程序','" & txtEdit(参数.入口程序).Text & "',3)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & TYPE_沈阳市 & ",NULL,'登录操作员','" & txtEdit(参数.登录操作员).Text & "',4)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & TYPE_沈阳市 & ",NULL,'登录口令','" & txtEdit(参数.登录口令).Text & "',5)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & TYPE_沈阳市 & ",NULL,'个人帐户支出(挂号)','" & cbo收入项目.ItemData(cbo收入项目.ListIndex) & "',6)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & TYPE_沈阳市 & ",NULL,'允许门诊业务','" & chk是否允许在院病人进行门诊业务.Value & "',7)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
'    gstrSQL = "zl_保险参数_Insert(" & TYPE_沈阳市 & ",NULL,'允许急救用药','" & chk急救用药.Value & "',8)"
'    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & TYPE_沈阳市 & ",NULL,'适用地区','" & cbo适用地区.ListIndex & "',9)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    gcnOracle.CommitTrans
    mblnReturn = True
    
    Call SaveSetting("ZLSOFT", "公共模块\" & App.ProductName, "IC设备端口", txtEdit(参数.IC设备端口).Text)
    Unload Me
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
    gcnOracle.RollbackTrans
End Sub

Private Function Valid() As Boolean
    Dim strPara As String, arrPara
    Dim intDO As Integer, intUbound As Integer
    '检查是否必需输入的参数都输入了
    strPara = "服务器地址||服务器端口号||服务器入口程序||IC设备端口号||登录操作员||口令"
    arrPara = Split(strPara, "||")
    
    intUbound = txtEdit.Count - 1
    For intDO = 0 To intUbound
        If Trim(txtEdit(intDO)) = "" Then
            MsgBox arrPara(intDO) & "不能为空！", vbInformation, gstrSysName
            txtEdit(intDO).SetFocus
            Exit Function
        End If
    Next
    
    Valid = True
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim intDO As Integer, blnFind As Boolean
    Dim rsTemp As New ADODB.Recordset
    
    '获取收入项目
    gstrSQL = "Select ID,名称 From 收入项目 Order by 编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取收入项目")
    
    cbo收入项目.Clear
    cbo收入项目.AddItem ""
    Do While Not rsTemp.EOF
        cbo收入项目.AddItem Nvl(rsTemp!名称)
        cbo收入项目.ItemData(cbo收入项目.NewIndex) = rsTemp!ID
        rsTemp.MoveNext
    Loop
    cbo收入项目.ListIndex = 0
    
    With cbo适用地区
        .Clear
        .AddItem "其他地区"
        .AddItem "长春地区"
        .AddItem "沈阳地区"
        .ListIndex = 0
    End With
    
    '获取服务器地址、端口及入口名称('服务器地址','服务器端口号','服务器入口程序')
    gstrSQL = " Select 参数名,参数值 From 保险参数" & _
              " Where 险类=[1] And 参数名 Like '服务器%'"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取服务器地址、端口及入口名称", TYPE_沈阳市)
    
    With rsTemp
        Do While Not .EOF
            Select Case !参数名
            Case "服务器地址"
                txtEdit(参数.地址).Text = Nvl(!参数值)
            Case "服务器端口号"
                txtEdit(参数.端口号).Text = Nvl(!参数值)
            Case "服务器入口程序"
                txtEdit(参数.入口程序).Text = Nvl(!参数值)
            Case "登录操作员"
                txtEdit(参数.登录操作员).Text = Nvl(!参数值)
            Case "登录口令"
                txtEdit(参数.登录口令).Text = Nvl(!参数值)
            Case "个人帐户支出(挂号)"
                For intDO = 1 To cbo收入项目.ListCount
                    cbo收入项目.ListIndex = intDO - 1
                    If cbo收入项目.ItemData(cbo收入项目.ListIndex) = Nvl(!参数值, 0) Then
                        blnFind = True
                        Exit For
                    End If
                    If Not blnFind Then cbo收入项目.ListIndex = 0
                Next
            Case "允许门诊业务"
                chk是否允许在院病人进行门诊业务.Value = Nvl(!参数值, 0)
'            Case "允许急救用药"
'                chk急救用药.Value = NVL(!参数值, 0)
            Case "适用地区"
                cbo适用地区.ListIndex = Nvl(!参数值, 0)
            End Select
            .MoveNext
        Loop
    End With
    txtEdit(参数.IC设备端口).Text = GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "IC设备端口", 1)
    
    If cbo收入项目.ItemData(cbo收入项目.ListIndex) <> 0 Then chk挂号.Value = 1
End Sub

Public Function ShowME() As Boolean
    mblnReturn = False
    Me.Show 1
    ShowME = mblnReturn
End Function
