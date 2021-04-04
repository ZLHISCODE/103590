VERSION 5.00
Begin VB.Form frmIdentify贵阳补充结算 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "补充结算设置"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8520
   Icon            =   "frmIdentify贵阳补充结算.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   8520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdOK 
      Caption         =   "冲销(&E)"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5445
      TabIndex        =   26
      Top             =   4335
      Width           =   1335
   End
   Begin VB.Frame frmDetail 
      Height          =   4035
      Left            =   195
      TabIndex        =   30
      Top             =   150
      Width           =   8070
      Begin VB.TextBox txt医保号 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1515
         TabIndex        =   3
         Top             =   960
         Width           =   2160
      End
      Begin VB.TextBox txt主页ID 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   5385
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1485
         Width           =   2445
      End
      Begin VB.TextBox txt病人ID 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1515
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1440
         Width           =   2445
      End
      Begin VB.TextBox txt住院号 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   5385
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   1965
         Width           =   2445
      End
      Begin VB.TextBox txt姓名 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1515
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1905
         Width           =   2445
      End
      Begin VB.TextBox txt卡号 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   5385
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   960
         Width           =   2160
      End
      Begin VB.OptionButton opt门诊 
         Caption         =   "门诊"
         Height          =   375
         Left            =   1515
         TabIndex        =   0
         Top             =   285
         Value           =   -1  'True
         Width           =   810
      End
      Begin VB.OptionButton opt住院 
         Caption         =   "住院"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5385
         TabIndex        =   1
         Top             =   285
         Width           =   810
      End
      Begin VB.CommandButton cmd卡号 
         Height          =   300
         Left            =   7530
         Picture         =   "frmIdentify贵阳补充结算.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   960
         Width           =   300
      End
      Begin VB.CommandButton cmd医保号 
         Height          =   300
         Left            =   3660
         Picture         =   "frmIdentify贵阳补充结算.frx":00EA
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   960
         Width           =   300
      End
      Begin VB.TextBox txt就诊顺序号 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1515
         TabIndex        =   21
         Top             =   3060
         Width           =   2445
      End
      Begin VB.TextBox txt结算编号 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   5385
         TabIndex        =   23
         Top             =   3075
         Width           =   2445
      End
      Begin VB.ComboBox cbo支付类别 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1515
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   3540
         Width           =   2445
      End
      Begin VB.TextBox txt性别 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1515
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   2430
         Width           =   2445
      End
      Begin VB.TextBox txt身份证号 
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   5385
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   2430
         Width           =   2445
      End
      Begin VB.Label lab医保号 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "医保号(&Y)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   495
         TabIndex        =   2
         Top             =   1005
         Width           =   945
      End
      Begin VB.Label lab主页ID 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "主页ID(&P)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   4320
         TabIndex        =   10
         Top             =   1530
         Width           =   945
      End
      Begin VB.Label lab病人ID 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "病人ID(&S)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   495
         TabIndex        =   8
         Top             =   1485
         Width           =   945
      End
      Begin VB.Label lab卡号 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "卡号(&I)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   4530
         TabIndex        =   5
         Top             =   1005
         Width           =   735
      End
      Begin VB.Label lab住院号 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "住院号(&H)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   210
         Left            =   4320
         TabIndex        =   14
         Top             =   2010
         Width           =   945
      End
      Begin VB.Label lab姓名 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "姓名(&N)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   705
         TabIndex        =   12
         Top             =   1950
         Width           =   735
      End
      Begin VB.Line Line3 
         BorderColor     =   &H000000FF&
         X1              =   0
         X2              =   60000
         Y1              =   2895
         Y2              =   2895
      End
      Begin VB.Line Line2 
         BorderColor     =   &H0080FFFF&
         X1              =   0
         X2              =   60000
         Y1              =   2910
         Y2              =   2910
      End
      Begin VB.Line Line1 
         BorderColor     =   &H0080FFFF&
         X1              =   0
         X2              =   60000
         Y1              =   855
         Y2              =   855
      End
      Begin VB.Line Line4 
         BorderColor     =   &H000000FF&
         X1              =   0
         X2              =   60000
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label lab就诊顺序号 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "就诊顺序号(&J)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   75
         TabIndex        =   20
         Top             =   3105
         Width           =   1365
      End
      Begin VB.Label lab结算编号 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "结算编号(&B)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   4125
         TabIndex        =   22
         Top             =   3120
         Width           =   1155
      End
      Begin VB.Label lbl支付类别 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "支付类别(&T)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   285
         TabIndex        =   24
         Top             =   3600
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "性别(&X)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   705
         TabIndex        =   16
         Top             =   2475
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "身份证号(&F)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   210
         Left            =   4110
         TabIndex        =   18
         Top             =   2475
         Width           =   1155
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "关闭(&C)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6930
      TabIndex        =   27
      Top             =   4335
      Width           =   1335
   End
   Begin VB.PictureBox P2 
      Height          =   495
      Left            =   1530
      Picture         =   "frmIdentify贵阳补充结算.frx":014A
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   29
      Top             =   7035
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox P1 
      Height          =   495
      Left            =   165
      Picture         =   "frmIdentify贵阳补充结算.frx":0228
      ScaleHeight     =   435
      ScaleWidth      =   1155
      TabIndex        =   28
      Top             =   7035
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frmIdentify贵阳补充结算"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mintInsure      As Integer

Private Sub cmdCancel_Click()
    With g补充结算
        .blnYn = False
    End With
    Unload Me
End Sub

Private Sub cmdOK_Click()
On Error GoTo ErrH
    Dim rsTmp                   As ADODB.Recordset
    Dim str就诊顺序号           As String
    Dim str结算编号             As String
    Dim str支付类型             As String
    Dim lng病人ID               As String
    Dim strMsg                  As String
    
    str就诊顺序号 = txt就诊顺序号.Text
    str结算编号 = txt结算编号.Text
    lng病人ID = Val(txt病人ID.Text)
    str支付类型 = cbo支付类别.ItemData(cbo支付类别.ListIndex)
    '检测数据在本地是否存在 如果存在则不能删除
    gstrSQL = "Select count(*) as Cnt From 保险结算记录 where 支付顺序号=[1] and 备注=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, str结算编号, IIf(cbo支付类别.Text = "普通门诊", "普通", "特殊") & str就诊顺序号)
    If rsTmp!cnt > 0 Then
        strMsg = "你所要冲销的单据信息" & vbCrLf
        strMsg = strMsg & "就诊顺序号：【" & str就诊顺序号 & "】" & vbCrLf
        strMsg = strMsg & "结算编号：【" & txt结算编号 & "】" & vbCrLf
        strMsg = strMsg & "支付类型：【" & cbo支付类别.Text & "】" & vbCrLf
        strMsg = strMsg & "在HIS系统中已存在，请到HIS收费系统中进行冲销！"
        MsgBox strMsg, vbCritical, gstrSysName
        Exit Sub
    End If
    strMsg = "是否冲销单据？需要冲销的单据信息如下：" & vbCrLf
    strMsg = strMsg & "就诊顺序号：【" & str就诊顺序号 & "】" & vbCrLf
    strMsg = strMsg & "结算编号：【" & txt结算编号 & "】" & vbCrLf
    strMsg = strMsg & "支付类型：【" & cbo支付类别.Text & "】" & vbCrLf
    strMsg = strMsg & "选择[是]将冲销中心的单据！"
    If MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    '对XML DomDocument对象进行初始化
    If InitXML = False Then Exit Sub
    Call InsertChild(mdomInput.documentElement, "BILLNO", str就诊顺序号)     ' 就诊顺序号
    Call InsertChild(mdomInput.documentElement, "BALANCEID", str结算编号)    ' 结算编号
    Call InsertChild(mdomInput.documentElement, "PAYTYPE", str支付类型)     ' 支付类别
    Call InsertChild(mdomInput.documentElement, "OPERATOR", UserInfo.姓名)     ' 操作员
    Call InsertChild(mdomInput.documentElement, "DODATE", Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss"))  ' 办理日期
    
    '调用接口
    If MsgBox("请再次确认是否退票？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    If CommServer("RETBALANCE", IIf(IS离休(lng病人ID), 1, 0)) Then
        '保存冲销信息
        gstrSQL = "ZL_冲销中心收费_Update('" & str就诊顺序号 & "','" & str结算编号 & "','" & str支付类型 & "—" & cbo支付类别.Text & "','" & txt医保号.Text & "','" & txt卡号.Text & "','" & txt病人ID.Text & "','" & IIf(opt门诊.Value, 0, txt主页ID.Text) & "','" & txt姓名.Text & "','" & txt性别.Text & "','" & txt身份证号.Text & "','" & IIf(opt门诊.Value, 0, txt住院号.Text) & "','" & UserInfo.姓名 & "','" & zlDatabase.Currentdate & "' ,'成功')"
        MsgBox "冲销中心数据成功！", vbExclamation, gstrSysName
    Else
        gstrSQL = "ZL_冲销中心收费_Update('" & str就诊顺序号 & "','" & str结算编号 & "','" & str支付类型 & "—" & cbo支付类别.Text & "','" & txt医保号.Text & "','" & txt卡号.Text & "','" & txt病人ID.Text & "','" & IIf(opt门诊.Value, 0, txt主页ID.Text) & "','" & txt姓名.Text & "','" & txt性别.Text & "','" & txt身份证号.Text & "','" & IIf(opt门诊.Value, 0, txt住院号.Text) & "','" & UserInfo.姓名 & "','" & zlDatabase.Currentdate & "' ,'失败')"
        MsgBox "冲销中心数据失败！", vbExclamation, gstrSysName
    End If
    Call zlDatabase.ExecuteProcedure(gstrSQL, "冲销中心收费")
    Exit Sub
ErrH:
    MsgBox "冲销中心数据失败！" & vbCrLf & Err.Description, vbCritical, gstrSysName
    Err.Clear
    Exit Sub
End Sub

Private Sub cmd卡号_Click()
    
    Dim rsTmp       As ADODB.Recordset
    
    cmd卡号.Picture = P2.Picture
    txt卡号.Locked = False
    txt卡号.ForeColor = vbBlue
    txt卡号.BackColor = &HC0FFC0
    txt卡号.SetFocus
    txt卡号.SelStart = 0
    txt卡号.SelLength = Len(txt卡号.Text)
    cmd医保号.Picture = P1.Picture
    txt医保号.Locked = True
    txt医保号.ForeColor = &HFF00FF
    txt医保号.BackColor = vbWhite
    
    If Trim(txt卡号.Text) = "" Then Exit Sub
    
    gstrSQL = "select * from 病人信息 where 险类=[1] And IC卡号=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mintInsure, txt卡号.Text)
    If rsTmp.RecordCount = 1 Then
        txt卡号.Text = "" & rsTmp!IC卡号
        txt医保号.Text = "" & rsTmp!医保号
        txt病人ID.Text = "" & rsTmp!病人ID
        txt主页ID = 0
        txt姓名.Text = rsTmp!姓名
        txt住院号.Text = "" & rsTmp!住院号
        txt性别.Text = "" & rsTmp!性别
        txt身份证号.Text = "" & rsTmp!身份证号
        cmdOK.Enabled = True
    Else
        MsgBox "末找到相关病人信息", vbCritical, gstrSysName
        cmdOK.Enabled = False
    End If
    
End Sub

Private Sub cmd医保号_Click()
    Dim rsTmp       As ADODB.Recordset
    
    cmd医保号.Picture = P2.Picture
    txt医保号.Locked = False
    txt医保号.ForeColor = vbBlue
    txt医保号.BackColor = &HC0FFC0
    txt医保号.SetFocus
    txt医保号.SelStart = 0
    txt医保号.SelLength = Len(txt医保号.Text)
    cmd卡号.Picture = P1.Picture
    txt卡号.Locked = True
    txt卡号.ForeColor = &HFF00FF
    txt卡号.BackColor = vbWhite
    
    If Trim(txt医保号) = "" Then Exit Sub
    
    gstrSQL = "select * from 病人信息 where 险类=[1] And 医保号=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mintInsure, txt医保号.Text)
    If rsTmp.RecordCount = 1 Then
        txt卡号.Text = "" & rsTmp!IC卡号
        txt医保号.Text = "" & rsTmp!医保号
        txt病人ID.Text = "" & rsTmp!病人ID
        txt主页ID = 0
        txt姓名.Text = rsTmp!姓名
        txt住院号.Text = "" & rsTmp!住院号
        txt性别.Text = "" & rsTmp!性别
        txt身份证号.Text = "" & rsTmp!身份证号
        cmdOK.Enabled = True
    Else
        MsgBox "末找到相关病人信息", vbCritical, gstrSysName
        cmdOK.Enabled = False
    End If
End Sub

Private Sub Form_Activate()
    Call cmd医保号_Click
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    lab主页ID.ForeColor = &H80000003
    txt主页ID.Enabled = False
    lab住院号.ForeColor = &H80000003
    txt住院号.Enabled = False
    With cbo支付类别
        .AddItem "普通门诊"
        .ItemData(.NewIndex) = 11
        .AddItem "特殊门诊"
        .ItemData(.NewIndex) = 18
        .ListIndex = 0
    End With
End Sub

Private Sub opt门诊_Click()
    Call ControlRefech
End Sub

Private Sub opt住院_Click()
    Call ControlRefech
End Sub

Private Sub txt病人ID_Validate(Cancel As Boolean)
    If Val(txt病人ID.Text) <= 0 Then Cancel = True Else txt病人ID.Text = Val(txt病人ID.Text)
End Sub

Private Sub txt主页ID_Validate(Cancel As Boolean)
    If Val(txt主页ID.Text) <= 0 Then Cancel = True Else txt主页ID.Text = Val(txt主页ID.Text)
End Sub

Public Sub ControlRefech()
On Error GoTo ErrH
    If opt门诊.Value Then
        lab主页ID.ForeColor = &H80000003
        txt主页ID.Enabled = False
        lab住院号.ForeColor = &H80000003
        txt住院号.Enabled = False
        
        With cbo支付类别
            .Clear
            .AddItem "普通门诊"
            .ItemData(.NewIndex) = 11
            .AddItem "特殊门诊"
            .ItemData(.NewIndex) = 18
            .ListIndex = 0
        End With
    Else
        lab主页ID.ForeColor = vbBlack
        txt主页ID.Enabled = True
        lab住院号.ForeColor = vbBack
        txt住院号.Enabled = True
        With cbo支付类别
            .Clear
            .AddItem "普通住院"
            .ItemData(.NewIndex) = 31
            .ListIndex = 0
        End With
    End If
    Exit Sub
ErrH:
    Err.Clear
    Exit Sub
End Sub
 

Public Property Let Insure(ByVal vNewValue As Integer)
    mintInsure = vNewValue
End Property

