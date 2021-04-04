VERSION 5.00
Begin VB.Form frm保险结账上传_贵阳编辑 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "贵阳市医保单个病人上传"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6585
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   10.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5983.333
   ScaleMode       =   0  'User
   ScaleWidth      =   6585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame fra生育备案编辑 
      Caption         =   "单个病人上传"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4500
      Left            =   323
      TabIndex        =   2
      Top             =   210
      Width           =   5715
      Begin VB.TextBox Txt记录id 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   315
         Left            =   2520
         MaxLength       =   20
         TabIndex        =   25
         Top             =   4080
         Width           =   2775
      End
      Begin VB.TextBox txt收费时间 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2505
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   3660
         Width           =   2775
      End
      Begin VB.TextBox txt结束日期 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2505
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   3285
         Width           =   2775
      End
      Begin VB.TextBox txt姓名 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2505
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1740
         Width           =   2775
      End
      Begin VB.TextBox txt年龄 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2505
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   2520
         Width           =   2775
      End
      Begin VB.TextBox txt性别 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2505
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   2130
         Width           =   2775
      End
      Begin VB.TextBox txt住院号 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Height          =   315
         Left            =   2505
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   600
         Width           =   2775
      End
      Begin VB.TextBox txt就诊顺序号 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   2505
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1350
         Width           =   2775
      End
      Begin VB.TextBox txt开始日期 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2505
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   2910
         Width           =   2775
      End
      Begin VB.TextBox txt支付类别 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   2505
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   975
         Width           =   2775
      End
      Begin VB.CommandButton Cmd病人ID 
         Caption         =   "…"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4980
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   210
         Width           =   285
      End
      Begin VB.TextBox txt医保号 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2505
         MaxLength       =   14
         TabIndex        =   8
         Top             =   232
         Width           =   2775
      End
      Begin VB.Label Lab记录id 
         Caption         =   "记录id"
         Height          =   255
         Left            =   1200
         TabIndex        =   24
         Top             =   4080
         Width           =   735
      End
      Begin VB.Label lab收费时间 
         AutoSize        =   -1  'True
         Caption         =   "收费时间"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1215
         TabIndex        =   23
         Top             =   3720
         Width           =   780
      End
      Begin VB.Label lab结束日期 
         AutoSize        =   -1  'True
         Caption         =   "结束日期"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1215
         TabIndex        =   22
         Top             =   3345
         Width           =   780
      End
      Begin VB.Label lab姓名 
         AutoSize        =   -1  'True
         Caption         =   "姓名"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1605
         TabIndex        =   19
         Top             =   1800
         Width           =   390
      End
      Begin VB.Label lab性别 
         AutoSize        =   -1  'True
         Caption         =   "性别"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1605
         TabIndex        =   18
         Top             =   2190
         Width           =   390
      End
      Begin VB.Label lab年龄 
         AutoSize        =   -1  'True
         Caption         =   "年龄"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1605
         TabIndex        =   17
         Top             =   2580
         Width           =   390
      End
      Begin VB.Label lab住院号 
         AutoSize        =   -1  'True
         Caption         =   "住院号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1410
         TabIndex        =   13
         Top             =   660
         Width           =   585
      End
      Begin VB.Label lab开始日期 
         AutoSize        =   -1  'True
         Caption         =   "开始日期"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1215
         TabIndex        =   12
         Top             =   2970
         Width           =   780
      End
      Begin VB.Label lab就诊顺序号 
         AutoSize        =   -1  'True
         Caption         =   "就诊顺序号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1020
         TabIndex        =   11
         Top             =   1410
         Width           =   975
      End
      Begin VB.Label lab支付类别 
         AutoSize        =   -1  'True
         Caption         =   "支付类别"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1215
         TabIndex        =   10
         Top             =   1035
         Width           =   780
      End
      Begin VB.Label lab医保号 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "医保号(&B)*"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1125
         TabIndex        =   9
         Top             =   285
         Width           =   1005
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   4845
      TabIndex        =   1
      Top             =   4785
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "保存并上传(&O)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   2805
      TabIndex        =   0
      Top             =   4785
      Width           =   1500
   End
   Begin VB.Shape sapStatus 
      BorderColor     =   &H000000FF&
      BorderStyle     =   6  'Inside Solid
      DrawMode        =   9  'Not Mask Pen
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   315
      Left            =   1800
      Shape           =   3  'Circle
      Top             =   4800
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "frm保险结账上传_贵阳编辑"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstr生育住院            As String
Private mintInsure              As Integer
Private mblnOkCancel            As Boolean
Private mstr就诊顺序号          As String
Dim sngX                        As Single
Dim sngY                        As Single
Dim sngH                        As Single
Dim strUp                       As String

Public Property Let Insure(ByVal vNewValue As Integer)
    mintInsure = vNewValue
End Property

Public Property Get OkCancel() As Boolean
    OkCancel = mblnOkCancel
End Property

Public Property Get 就诊顺序号() As String
    就诊顺序号 = mstr就诊顺序号
End Property

Public Property Let 就诊顺序号(ByVal vNewValue As String)
    mstr就诊顺序号 = vNewValue
End Property

Private Sub cmdADD_Click()
    Dim strDateTime As String
On Error GoTo ErrH
    mstr就诊顺序号 = txt就诊顺序号.Text
    If Not mCheckData Then Exit Sub
     
    '对XML DomDocument对象进行初始化
    If InitXML = False Then Exit Sub
    Call InsertChild(mdomInput.documentElement, "PAYTYPE", txt支付类别.Text)    ' 支付类别
    Call InsertChild(mdomInput.documentElement, "BILLNO", mstr就诊顺序号)       ' 就诊顺序号
    '调用接口
    If CommServer("UPLOADBYBILLNO") = False Then Exit Sub

    '保存
    gstrSQL = "Zl_贵阳_结算上传_Update('" & txt支付类别.Text & "','" & mstr就诊顺序号 & "','" & txt医保号.Tag & "','" & txt住院号.Text & "','" & txt姓名.Text & "','" & txt性别.Text & "','" & txt年龄.Text & "'," & IIf(IsDate(txt开始日期.Text), "to_date('" & txt开始日期.Text & "','yyyy-mm-dd hh24:mi:ss')", "Null") & "," & IIf(IsDate(txt结束日期.Text), "to_date('" & txt结束日期.Text & "','yyyy-mm-dd hh24:mi:ss')", "Null") & "," & IIf(IsDate(txt收费时间.Text), "to_date('" & txt收费时间.Text & "','yyyy-mm-dd hh24:mi:ss')", "Null") & ",'" & UserInfo.姓名 & "',sysdate,'" & Txt记录id.Text & "')"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
'
    mblnOkCancel = True
    txt医保号.Text = ""
    txt医保号.Tag = ""
    txt姓名.Text = ""
    sapStatus.Visible = True
    sapStatus.FillColor = &HC000&
    sapStatus.BorderColor = &HC000&
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    sapStatus.FillColor = vbRed
    sapStatus.BorderColor = vbRed
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdCancel_Click()
On Error GoTo ErrH
    Unload Me
    Exit Sub
ErrH:
    Err.Clear
    Exit Sub
End Sub

Private Sub cmdOK_Click()
    Dim strDateTime     As String
On Error GoTo ErrH
    mstr就诊顺序号 = txt就诊顺序号.Text
    If Not mCheckData Then Exit Sub
     
    '对XML DomDocument对象进行初始化
    If InitXML = False Then Exit Sub
    Call InsertChild(mdomInput.documentElement, "PAYTYPE", txt支付类别.Text)    ' 支付类别
    Call InsertChild(mdomInput.documentElement, "BILLNO", mstr就诊顺序号)       ' 就诊顺序号
    '调用接口
    If CommServer("UPLOADBYBILLNO") = False Then Exit Sub

    '保存
    gstrSQL = "Zl_贵阳_结算上传_Update('" & txt支付类别.Text & "','" & mstr就诊顺序号 & "','" & txt医保号.Tag & "','" & txt住院号.Text & "','" & txt姓名.Text & "','" & txt性别.Text & "','" & txt年龄.Text & "'," & IIf(IsDate(txt开始日期.Text), "to_date('" & txt开始日期.Text & "','yyyy-mm-dd hh24:mi:ss')", "Null") & "," & IIf(IsDate(txt结束日期.Text), "to_date('" & txt结束日期.Text & "','yyyy-mm-dd hh24:mi:ss')", "Null") & "," & IIf(IsDate(txt收费时间.Text), "to_date('" & txt收费时间.Text & "','yyyy-mm-dd hh24:mi:ss')", "Null") & ",'" & UserInfo.姓名 & "',sysdate,'" & Txt记录id.Text & "')"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
'
    mblnOkCancel = True
    Unload Me
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Err.Clear
    Exit Sub
End Sub

Private Sub Cmd病人ID_Click()
    Dim rsTemp As New ADODB.Recordset
    Dim vRect   As RECT
    
    'gstrSQL = "select /*+ rule */C.就诊流水号 as ID,d.id as 记录id,B.医保号,B.住院号,C.医疗类别 As 支付类别,C.就诊流水号 as 就诊顺序号,B.姓名,B.性别,B.年龄, A.入院日期,A.出院日期,D.收费时间" & vbNewLine & _
       '         "from 病案主页 A,病人信息 B,保险结算记录 C,病人结帐记录 D" & vbNewLine & _
        '        "where A.出院日期>sysdate-[2] And C.性质=2 And A.病人ID=B.病人ID And A.病人ID=C.病人ID And C.记录id = D.Id And B.住院号 is not null And A.险类=[1]" & vbNewLine & _
        '   '     "And C.就诊流水号 Not In (Select 就诊顺序号 From 贵阳_结算上传)"
    
    gstrSQL = "Select /*+ rule */ Distinct c.就诊流水号 As Id, a.Id As 记录id, b.医保号, b.住院号,b.姓名 , b.性别, b.年龄, a.收费时间, c.支付类别 As 支付类别, c.就诊流水号 As 就诊顺序号," & vbNewLine & _
            "  a.开始日期, a.结束日期" & vbNewLine & _
            "From 病人结帐记录 a, 病人信息 b," & vbNewLine & _
           "(Select Distinct b.住院号, b.病人id, a.就诊流水号, a.记录id, a.医疗类别 As 支付类别, a.结算时间 " & vbNewLine & _
            "From 保险结算记录 a, 病案主页 b" & vbNewLine & _
            "Where a. 病人id = b.病人id And a.结算时间 >=sysdate-1000  And b.险类 = [1] And a.性质 = 2) c" & vbNewLine & _
            "Where a.病人id = c.病人id And a.Id = c.记录id And a.收费时间 >= Sysdate - 1000 And c.病人id = b.病人id And b.险类 = [1] And " & vbNewLine & _
           " a.记录状态 = 1 And a.收费时间 >= Sysdate -1000  And C.就诊流水号 Not In (Select 就诊顺序号 From 贵阳_结算上传)  " & vbNewLine & _
           " order by b.医保号"
    
    
    vRect = GetControlRect(txt医保号.hwnd)
    sngX = vRect.Left
    sngY = vRect.Top
    sngH = txt医保号.Height
 
    DoEvents
    Set rsTemp = zlDatabase.ShowSQLSelect( _
            Nothing, gstrSQL, 0, "结算信息查询", False, _
            "", "", False, False, True, _
            sngX, sngY, sngH, False, False, _
            False, mintInsure, 90 _
            )
    If ChkRsState(rsTemp) Then
        txt医保号.Tag = ""
        txt医保号.Text = ""
        txt住院号.Text = ""
        txt姓名.Text = ""
        txt性别.Text = ""
        txt年龄.Text = ""
        txt支付类别.Text = ""
        txt就诊顺序号.Text = ""
        Txt记录id.Text = ""
        txt开始日期.Text = ""
        txt结束日期.Text = ""
        txt收费时间.Text = ""
    Else
        txt医保号.Tag = Nvl(rsTemp!医保号)
        txt医保号.Text = Nvl(rsTemp!医保号)
        txt住院号.Text = Nvl(rsTemp!住院号)
        txt姓名.Text = Nvl(rsTemp!姓名)
        txt性别.Text = Nvl(rsTemp!性别)
        txt年龄.Text = Nvl(rsTemp!年龄)
        txt支付类别.Text = Nvl(rsTemp!支付类别)
        txt就诊顺序号.Text = Nvl(rsTemp!就诊顺序号)
        Txt记录id.Text = Nvl(rsTemp!记录ID)
        txt开始日期.Text = Format(Nvl(rsTemp!开始日期), "yyyy-mm-dd hh:mm:ss")
        txt结束日期.Text = Format(Nvl(rsTemp!结束日期), "yyyy-mm-dd hh:mm:ss")
        txt收费时间.Text = Format(Nvl(rsTemp!收费时间), "yyyy-mm-dd hh:mm:ss")
        
        zlCommFun.PressKey vbKeyTab
End If
End Sub

Private Sub txt性别_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txt姓名_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txt医保号_KeyPress(KeyAscii As Integer)
    Dim rsTemp      As New ADODB.Recordset
    Dim vRect       As RECT
    Dim strText     As String
    
    If KeyAscii <> 13 Then Exit Sub
    If Trim(txt医保号.Text) = "" Then Exit Sub
    If txt医保号.Locked Then
        zlCommFun.PressKey vbKeyTab
        Exit Sub
    End If
    strText = txt医保号.Text
  '  gstrSQL = "select /*+ rule */ C.就诊流水号 as ID,d.id as 记录id,B.医保号,B.住院号,C.医疗类别 As 支付类别,C.就诊流水号 as 就诊顺序号,B.姓名,B.性别,B.年龄, A.入院日期,A.出院日期,D.收费时间" & vbNewLine & _
   '             "from 病案主页 A,病人信息 B,保险结算记录 C,病人结帐记录 D" & vbNewLine & _
     '           "where A.出院日期>sysdate-[2] And C.性质=2 And A.病人ID=B.病人ID And A.病人ID=C.病人ID And C.记录id = D.Id And B.住院号 is not null And A.险类=[1]" & vbNewLine & _
      '          "And C.就诊流水号 Not In (Select 就诊顺序号 From 贵阳_结算上传)"
                
                
     gstrSQL = "Select /*+ rule */ distinct c.就诊流水号 As Id, a.Id As 记录id, b.医保号, b.住院号,b.姓名 , b.性别, b.年龄, a.收费时间, c.支付类别 As 支付类别, c.就诊流水号 As 就诊顺序号," & vbNewLine & _
            "  a.开始日期, a.结束日期" & vbNewLine & _
            "From 病人结帐记录 a, 病人信息 b," & vbNewLine & _
           "(Select Distinct b.住院号, b.病人id, a.就诊流水号, a.记录id, a.医疗类别 As 支付类别, a.结算时间 " & vbNewLine & _
            "From 保险结算记录 a, 病案主页 b" & vbNewLine & _
            "Where a. 病人id = b.病人id And a.结算时间 >=sysdate-1000  And b.险类 = [1] And a.性质 = 2) c" & vbNewLine & _
            "Where a.病人id = c.病人id And a.Id = c.记录id And a.收费时间 >= Sysdate - 1000 And c.病人id = b.病人id And b.险类 = [1] And " & vbNewLine & _
           " a.记录状态 = 1 And a.收费时间 >= Sysdate -1000  And C.就诊流水号 Not In (Select 就诊顺序号 From 贵阳_结算上传)"
           
    If zlCommFun.IsCharAlpha(strText) Then
        gstrSQL = gstrSQL & vbCrLf & "And zlspellcode(B.姓名) like '" & UCase(strText) & "%'"
    ElseIf zlCommFun.IsNumOrChar(strText) Then
        gstrSQL = gstrSQL & vbCrLf & "And (B.住院号 like '" & UCase(strText) & "%' or B.医保号 like  '" & UCase(strText) & "%')"
    ElseIf zlCommFun.IsCharChinese(strText) Then
        gstrSQL = gstrSQL & vbCrLf & "And B.姓名 like '" & UCase(strText) & "%'"
    Else
        gstrSQL = gstrSQL & vbCrLf & "And B.医保号 like '" & UCase(strText) & "%'"
    End If
    
    vRect = GetControlRect(txt医保号.hwnd)
    sngX = vRect.Left
    sngY = vRect.Top
    sngH = txt医保号.Height
    
    'DoEvents
    Set rsTemp = zlDatabase.ShowSQLSelect( _
            Nothing, gstrSQL, 0, "结算信息查询", False, _
            "", "", False, False, True, _
            sngX, sngY, sngH, False, False, _
            False, mintInsure, 90 _
            )
    If ChkRsState(rsTemp) Then
        txt医保号.Tag = ""
        txt医保号.Text = ""
        txt住院号.Text = ""
        txt姓名.Text = ""
        txt性别.Text = ""
        txt年龄.Text = ""
        txt支付类别.Text = ""
        txt就诊顺序号.Text = ""
        Txt记录id.Text = ""
        txt开始日期.Text = ""
        txt结束日期.Text = ""
        txt收费时间.Text = ""
    Else
        txt医保号.Tag = Nvl(rsTemp!医保号)
        txt医保号.Text = Nvl(rsTemp!医保号)
        txt住院号.Text = Nvl(rsTemp!住院号)
        txt姓名.Text = Nvl(rsTemp!姓名)
        txt性别.Text = Nvl(rsTemp!性别)
        txt年龄.Text = Nvl(rsTemp!年龄)
        txt支付类别.Text = Nvl(rsTemp!支付类别)
        txt就诊顺序号.Text = Nvl(rsTemp!就诊顺序号)
        Txt记录id.Text = Nvl(rsTemp!记录ID)
        txt开始日期.Text = Format(Nvl(rsTemp!开始日期), "yyyy-mm-dd hh:mm:ss")
        txt结束日期.Text = Format(Nvl(rsTemp!结束日期), "yyyy-mm-dd hh:mm:ss")
        txt收费时间.Text = Format(Nvl(rsTemp!收费时间), "yyyy-mm-dd hh:mm:ss")
        zlCommFun.PressKey vbKeyTab
End If
End Sub

Private Sub Form_Load()
    Dim rsTmp       As ADODB.Recordset
    Dim strDate     As String
    
    strDate = zlDatabase.Currentdate
    If mstr就诊顺序号 <> "" Then

'        gstrSQL = "Select /*+ rule */" & vbCrLf & _
'                "to_char(A.险类) || to_char(A.病种ID) || to_char(A.医保号) AS ID," & vbCrLf & _
'                "A.险类,A.病种ID,B.名称 AS 保险名称,C.编码 AS 病种编码,C.名称 as 病种名称," & vbCrLf & _
'                "A.医保号,A.姓名,A.性别,A.年龄," & vbCrLf & _
'                "A.备注, A.登记人,A.登记日期," & vbCrLf & _
'                "A.取消人 , A.取消日期, A.取消原因" & vbCrLf & _
'                "FROM 大连_特病人员 A,保险类别 B,保险病种 C" & vbCrLf & _
'                "Where A.险类 = B.序号 And A.险类 = C.险类 And A.病种ID = C.ID" & vbCrLf & _
'                "And A.险类=[1] and A.病种ID=[2] And A.医保号 =[3]"
'        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mintInsure, mstr就诊顺序号, mstr就诊顺序号)
'        If Not ChkRsState(rsTmp) Then
'            '修改状态
'            cmdAdd.Visible = False
'            Cmd病人ID.Enabled = False
'        End If
    End If
    txt姓名.BackColor = gconLockColor
    txt性别.BackColor = gconLockColor
End Sub

Private Function mCheckData() As Boolean
On Error GoTo ErrH
    
    If txt医保号.Tag = "" Then
        MsgBox "就诊顺序号不能为空！", vbCritical, gstrSysName
        Exit Function
    End If
    '检测当前电脑号是否存在数据，如果存在则不能新加
    gstrSQL = "Select count(1) From 贵阳_结算上传 where 就诊顺序号=[1]"
    If zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mstr就诊顺序号).Fields(0) > 0 Then
        MsgBox "就诊顺序号【" & mstr就诊顺序号 & "】已登记存在！" & vbCrLf & "不允许重新录入或编辑！", vbCritical, gstrSysName
        Exit Function
    End If
    
    If MsgBox("请确认你所录入的数据！上传后将不能修改！", vbOKCancel + vbDefaultButton2 + vbQuestion, gstrSysName) <> vbOK Then
        Exit Function
    End If
    mCheckData = True
    Exit Function
ErrH:
    MsgBox Err.Description, vbCritical, gstrSysName
    Err.Clear
    Exit Function
End Function




