VERSION 5.00
Begin VB.Form frmIdentify南京市 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "病保病人身份验证"
   ClientHeight    =   4395
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   4590
   Icon            =   "frmIdentify南京市.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmd子医保 
      Caption         =   "子医保(&S)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   180
      TabIndex        =   8
      Top             =   3840
      Width           =   1365
   End
   Begin VB.Frame fra基本 
      Caption         =   "医保病人基本信息"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3475
      Left            =   120
      TabIndex        =   10
      Top             =   240
      Width           =   4404
      Begin VB.TextBox txt年龄 
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
         Height          =   375
         Left            =   1920
         TabIndex        =   3
         Top             =   1830
         Width           =   1695
      End
      Begin VB.ComboBox cbo性别 
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
         Height          =   360
         ItemData        =   "frmIdentify南京市.frx":000C
         Left            =   1920
         List            =   "frmIdentify南京市.frx":001C
         TabIndex        =   2
         Top             =   1360
         Width           =   1695
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   1920
         TabIndex        =   6
         Top             =   2820
         Width           =   1692
      End
      Begin VB.TextBox Text1 
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
         Height          =   372
         Left            =   1920
         TabIndex        =   1
         Top             =   855
         Width           =   1692
      End
      Begin VB.CommandButton cmd病种信息 
         Caption         =   "…"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   3624
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   2295
         Width           =   372
      End
      Begin VB.TextBox txt门诊病种 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   1920
         TabIndex        =   4
         Top             =   2295
         Width           =   1692
      End
      Begin VB.TextBox txt姓名 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   1920
         TabIndex        =   0
         Top             =   360
         Width           =   1692
      End
      Begin VB.Label Label4 
         Caption         =   "年龄"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   16
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "性别"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   15
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "正确姓名"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   855
         TabIndex        =   14
         Top             =   2880
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "姓名"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1320
         TabIndex        =   12
         Top             =   915
         Width           =   480
      End
      Begin VB.Label lbl门诊病种 
         AutoSize        =   -1  'True
         Caption         =   "门诊病种(&F)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   480
         TabIndex        =   13
         Top             =   2370
         Width           =   1320
      End
      Begin VB.Label lbl姓名 
         AutoSize        =   -1  'True
         Caption         =   "标识号(&N)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   720
         TabIndex        =   11
         Top             =   420
         Width           =   1080
      End
   End
   Begin VB.CommandButton cmdCancle 
      Caption         =   "取消(&C)"
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
      Left            =   3432
      TabIndex        =   9
      Top             =   3840
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
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
      Left            =   2232
      TabIndex        =   7
      Top             =   3840
      Width           =   1100
   End
End
Attribute VB_Name = "frmIdentify南京市"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbytType As Byte
Private mstrIdentify As String
Private mlng病人ID As Long, mlng病种ID As Long
Private mstr病人姓名 As String
Private mstr病种编码 As String
Private mstr病种名称 As String
Private mstr医保号 As String
Private mstrsubInsure  As String '保存子医保的返回数据：保险序号|优惠类别|医保号|余额|停用

Private mintInsure As Integer, mstrReturn As String

Private Sub cbo性别_KeyPress(KeyAscii As Integer)
  If cbo性别.ListIndex = -1 Then cbo性别.ListIndex = 0
End Sub

Private Sub cmdCancle_Click()
    mstrIdentify = ""
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strSQL As String
    Dim strIdentify As String
    Dim strAddition As String
    Dim lngSequence As String
'    Dim str性别 As String, str出生日期 As String
    On Error GoTo errHandle
    
    '判断是否输入医保病人姓名
    If Trim(Text1.Text) = "" Then
        MsgBox "未提取到医保病人姓名", vbInformation, gstrSysName
        txt姓名.SetFocus
        Exit Sub
    End If
    
    mstr病人姓名 = Trim(Text1.Text)
    mstr医保号 = txt姓名.Text
    
    If Trim(txt门诊病种.Text) = "" Or txt门诊病种 <> mstr病种名称 Then
        MsgBox "门诊病种未录入或有误", vbInformation, gstrSysName
        txt门诊病种.SetFocus
        Exit Sub
    End If
    
    If InStr(1, Me.Tag, "|") <> 0 Then
'        str性别 = Split(Me.Tag, "|")(0)
'        str出生日期 = Split(Me.Tag, "|")(1)
    End If
    
    '此处无法取得卡号和医保号,所以暂时填入保险病种序列,以后得到卡号后再进行修改
    lngSequence = Right(String(20, "0") & Text1.Tag, 20)
'      strInfo='0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);
'      8中心;9.顺序号;10人员身份;11帐户余额;12当前状态;13病种ID;14在职(1,2,3);15退休证号;16年龄段;17灰度级
'      18帐户增加累计;19帐户支出累计;20进入统筹累计;21统筹报销累计;22住院次数累计;23就诊类别
'      24本次起付线;25起付线累计;26基本统筹限额
    
    strIdentify = strIdentify & ";"                                       '0卡号
    strIdentify = strIdentify & txt姓名.Text & ";"                  '1医保号（个人编号）
    strIdentify = strIdentify & ";"                                 '2密码
    strIdentify = strIdentify & Text1.Text & ";"                   '3姓名
    strIdentify = strIdentify & cbo性别.Text & ";"                                 '4性别
    strIdentify = strIdentify & IIf(Trim(txt年龄.Text) = "", Format(zlDatabase.Currentdate, "yyyy-mm-dd"), Get出生日期("", Val(txt年龄.Text))) & ";"                       '5出生日期
    strIdentify = strIdentify & ";"                                 '6身份证
    strIdentify = strIdentify & ";"                               '7.单位名称(编码)
    strAddition = "0;"                                          '8.中心代码
    strAddition = strAddition & ";"                               '9.顺序号
    strAddition = strAddition & ";"                            '10人员身份
    strAddition = strAddition & "10000;"                              '11帐户余额
    strAddition = strAddition & "0;"                            '12当前状态
    strAddition = strAddition & mlng病种ID & ";"                 '13病种ID
    strAddition = strAddition & "1;"                            '14在职(1,2,3)
    strAddition = strAddition & mstrsubInsure & ";"             '15退休证号
    strAddition = strAddition & ";"                             '16年龄段
    strAddition = strAddition & ";"                             '17灰度级
    strAddition = strAddition & ";"                             '18帐户增加累计
    strAddition = strAddition & ";"                            '19帐户支出累计
    strAddition = strAddition & "0;"                            '20进入统筹累计
    strAddition = strAddition & "0;"                            '21统筹报销累计
    strAddition = strAddition & "0;"                             '22住院次数累计
    strAddition = strAddition & ";"                             '23就诊类型
    
    mlng病人ID = BuildPatiInfo(0, strIdentify & strAddition, mlng病人ID, TYPE_南京市)
    '返回格式:中间插入病人ID
    If mlng病人ID > 0 Then
        mstrIdentify = strIdentify & mlng病人ID & ";" & strAddition
    End If
    If Trim(Text2.Text) <> "" Then
        mstr病人姓名 = Trim(Text2.Text)
    Else
        mstr病人姓名 = Trim(Text1.Text)
    End If
    
    Unload Me
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    MsgBox mlng病人ID & "，" & lngSequence & "，" & Text1.Tag & "，" & strIdentify & strAddition, vbInformation, gstrSysName
End Sub


Private Sub cmd病种信息_Click()
    Dim rsTemp As New ADODB.Recordset
    
    gstrSQL = "select id,编码,名称,decode(类别,1,'慢性病',2,'特殊病','普通病') as 病种 from 保险病种 where 险类=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "选择病种", TYPE_南京市)
    
    If frmListSel.ShowSelect(TYPE_南京市, rsTemp, "ID", "医保病种选择", "请选择特定的医保病种：") Then
        txt门诊病种.Text = rsTemp!名称
        mlng病种ID = rsTemp!ID
        mstr病种编码 = rsTemp!编码
        mstr病种名称 = rsTemp!名称
    Else
        txt门诊病种.SetFocus
    End If
End Sub

Private Sub cmd子医保_Click()
    '显示各子医保的身份验证窗体
    '返回数据:保险序号|优惠类别|医保号|余额|停用
    If Trim(txt姓名.Text) = "" Then
        MsgBox "请先确认医保病人身份！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    mstrsubInsure = frm子医保身份验证.ShowME(Me.Text1.Text)
    cmdOK.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub Form_Load()
    Dim str保险类别 As String
    If mbytType = 0 Or mbytType = 3 Then
        txt门诊病种.Enabled = True
    Else
        txt门诊病种.Enabled = False
    End If
    
    str保险类别 = GetSetting("ZLSOFT", "公共全局", "下属医保接口", "")
    cmd子医保.Enabled = (str保险类别 <> "")
End Sub

Private Sub txt门诊病种_GotFocus()
    OpenIme ("")
    Call zlControl.TxtSelAll(txt门诊病种)
End Sub

Private Sub txt门诊病种_KeyPress(KeyAscii As Integer)
    Dim rsTemp As New ADODB.Recordset
    Dim strText As String
    Dim blnReturn As Boolean
    
    If KeyAscii <> vbKeyReturn Then Exit Sub
    
    On Error GoTo errorhandle
    '读出门诊病种

    strText = txt门诊病种.Text
    gstrSQL = "select A.id,A.编码,A.名称 from 保险病种 A where A.险类=[1] and (A.编码 like [2] || '%' or A.名称 like [2] || '%' or A.简码 like [2] || '%')"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "门诊病种", TYPE_南京市, strText)
    
    If rsTemp.RecordCount = 1 Then
        blnReturn = True
    Else
        blnReturn = frmListSel.ShowSelect(TYPE_南京市, rsTemp, "ID", "医保病种选择", "请选择特定的医保病种：")
    End If
    
    If blnReturn Then
        txt门诊病种.Text = rsTemp!名称
        mlng病种ID = rsTemp!ID
        mstr病种编码 = rsTemp!编码
        mstr病种名称 = rsTemp!名称
        zlCommFun.PressKey (vbKeyTab)
    Else
        txt门诊病种_GotFocus
    End If
    Exit Sub
    
errorhandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub txt年龄_KeyPress(KeyAscii As Integer)
   If InStr("1234567890" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt姓名_GotFocus()
    Call zlControl.TxtSelAll(txt姓名)
End Sub

Public Function Identify(ByVal bytType As Byte, lng病人ID As Long) As String
    mbytType = bytType
    mlng病人ID = lng病人ID
    Me.Show 1
    Identify = mstrIdentify
    With gPatInfo_南京市
        .医保号 = mstr医保号
        .病人姓名 = mstr病人姓名
        .病种编码 = mstr病种编码
        .病种名称 = mstr病种名称
    End With
    lng病人ID = mlng病人ID
End Function

'Private Sub Txt姓名_KeyDown(KeyCode As Integer, Shift As Integer)
'    If Trim(txt姓名.Text) = "" Then KeyCode = 0
'    If KeyCode <> vbKeyReturn Then Exit Sub
'Dim rsTemp As ADODB.Recordset
'    gstrSQL = "select A.医保号,B.姓名,B.性别,B.年龄,B.工作单位" & _
'                " from 保险帐户 A,病人信息 B" & _
'                " Where A.医保号=[1] AND A.险类=[2] and A.病人ID=B.病人ID "
'    Set rsTemp = OpenSQLRecord(gstrSQL, "提取病人信息", CStr(txt姓名.Text), CInt(mintInsure))
'    If rsTemp.EOF Then '没在本医院就诊过
'        Text1.Locked = False
'        cbo性别.Locked = False
'        txt年龄.Locked = False
'
'    Else
'        txt姓名.Tag = Nvl(rsTemp!医保号)
'        Text1.Text = Nvl(rsTemp!姓名)
'        cbo性别.Text = Nvl(rsTemp!性别)
'        txt年龄.Text = Nvl(rsTemp!年龄)
'
'
'    End If
'
'End Sub

Private Sub txt姓名_KeyPress(KeyAscii As Integer)
    Dim datCurr As Date
    Dim StrInput As String, strSQL As String, rsTemp As New ADODB.Recordset
    
    
    On Error GoTo errHandle
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If txt姓名.Tag = "." Then Exit Sub
    
        If txt姓名 <> "" Then
        Me.Tag = ""
            If Left(txt姓名.Text, 1) = "." Then
        txt姓名.Tag = "."
        StrInput = txt姓名.Text
        If Not IsNumeric(Mid(StrInput, 2)) Then Exit Sub
        If Len(Mid(StrInput, 2)) <= 4 Then
            datCurr = zlDatabase.Currentdate()
            strSQL = PreFixNO & Format(CDate(Format(datCurr, "YYYY-MM-dd")) - CDate(Format(datCurr, "YYYY") & "-01-01") + 1, "000") & Format(Mid(StrInput, 2), "0000") '按天顺序编号
        Else
            strSQL = GetFullNO(Mid(StrInput, 2))
        End If
        '门诊记帐时必须要挂号建档
        strSQL = "Select A.病人id,A.姓名,A.标识号,A.性别,A.年龄 From 门诊费用记录 A Where A.NO='" & strSQL & "' And A.记录性质=4 And A.记录状态=1 "
        Set rsTemp = gcnOracle.Execute(strSQL)
'        If rsTemp.EOF Then
'            MsgBox "错误的挂号单号", vbInformation, gstrSysName
'            Exit Sub
'        End If
        mlng病人ID = rsTemp!病人ID
        strSQL = "Select a.姓名,a.门诊号,A.性别,A.年龄,b.医保号,c.ID,c.名称,c.编码 From 病人信息 a,保险帐户 b,保险病种 c Where a.病人id=b.病人id(+) and b.病种id=c.id(+) and a.病人ID=" & rsTemp!病人ID
        Set rsTemp = gcnOracle.Execute(strSQL)
        If rsTemp.EOF Then
            MsgBox "读取病人信息出错", vbInformation, gstrSysName
            Exit Sub
        ElseIf IsNull(rsTemp!门诊号) Then
            MsgBox "该病人的医保卡号没有录入", vbInformation, gstrSysName
            Exit Sub
        Else
            If IsNull(rsTemp!医保号) Then
                MsgBox "请补录该病人医保号", vbInformation, gstrSysName
            End If
            Text1.Text = rsTemp!姓名
            Text1.Tag = rsTemp!门诊号
            cbo性别.Text = rsTemp!性别
            txt年龄.Text = rsTemp!年龄
            txt门诊病种.Text = Nvl(rsTemp!名称)
            txt姓名.Text = Nvl(rsTemp!医保号)
            mlng病种ID = Nvl(rsTemp!ID, 0)
            mstr病种编码 = Nvl(rsTemp!编码)
            mstr病种名称 = Nvl(rsTemp!名称)
            '如果已存在该病人，将病人性别，年龄提取出来
'            Me.Tag = Nvl(rsTemp!性别) & "|" & Format(Nvl(rsTemp!出生日期, zlDatabase.Currentdate), "yyyy-MM-dd")
        End If
        Else
            txt姓名.Tag = ""
'            If mbytType = 0 Then
'            MsgBox "请输入市医保病人挂号单号", vbInformation, gstrSysName
'            Else
            Dim a As String
            a = txt姓名.Text
            Dim rsTemp1 As ADODB.Recordset
            gstrSQL = "select A.医保号,B.姓名,B.性别,B.年龄,c.ID,C.编码,c.名称" & _
                " from 保险帐户 A,病人信息 B,保险病种 c" & _
                " Where A.医保号=[1] AND A.险类=[2] and A.病人ID=B.病人ID and a.病种id=c.id(+)"
            Set rsTemp1 = zlDatabase.OpenSQLRecord(gstrSQL, "提取病人信息", CStr(txt姓名.Text), CInt(mintInsure))
            If rsTemp1.EOF Then '没在本医院就诊过
            Text1.Text = ""
            cbo性别.Text = ""
            txt年龄.Text = ""
            txt门诊病种.Text = ""
            Text1.Enabled = True
            cbo性别.Enabled = True
            txt年龄.Enabled = True
            Else
            txt姓名.Tag = Nvl(rsTemp1!医保号)
            Text1.Text = Nvl(rsTemp1!姓名)
            cbo性别.Text = Nvl(rsTemp1!性别)
            txt年龄.Text = Nvl(rsTemp1!年龄)
            txt门诊病种.Text = Nvl(rsTemp1!名称)
            mlng病种ID = Nvl(rsTemp1!ID, 0)
            mstr病种编码 = Nvl(rsTemp1!编码)
            mstr病种名称 = Nvl(rsTemp1!名称)
            End If
'            End If
'        Else
'          Text1.Locked = False
'          cbo性别.Locked = False
'          txt年龄.Locked = False
        End If
        Else
        zlCommFun.PressKey (vbKeyTab)
End If
    Exit Sub
errHandle:
    MsgBox "此医保病人没有建立档案", vbInformation, gstrSysName
End Sub


Private Sub txt姓名_Validate(Cancel As Boolean)
  txt姓名.Tag = Trim(txt姓名.Text)
End Sub


