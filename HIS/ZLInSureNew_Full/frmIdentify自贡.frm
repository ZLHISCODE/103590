VERSION 5.00
Begin VB.Form frmIdentify自贡 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "身份验证"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6855
   Icon            =   "frmIdentify自贡.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txt密码 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1770
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   930
      Width           =   1665
   End
   Begin VB.CheckBox chk离休人员 
      Caption         =   "离休人员"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   2445
   End
   Begin VB.CommandButton cmd修改密码 
      Caption         =   "改密码(&M)"
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
      Left            =   5280
      TabIndex        =   27
      Top             =   4440
      Width           =   1305
   End
   Begin VB.CommandButton cmd取消 
      Cancel          =   -1  'True
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
      Left            =   5310
      TabIndex        =   26
      Top             =   1140
      Width           =   1305
   End
   Begin VB.CommandButton cmd确认 
      Caption         =   "确认(&O)"
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
      Left            =   5310
      TabIndex        =   25
      Top             =   570
      Width           =   1305
   End
   Begin VB.Frame Frame1 
      Height          =   6405
      Left            =   4950
      TabIndex        =   28
      Top             =   -240
      Width           =   30
   End
   Begin VB.TextBox txt医保中心 
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
      Height          =   345
      Left            =   1770
      TabIndex        =   24
      Top             =   4590
      Width           =   2895
   End
   Begin VB.TextBox txt帐户余额 
      Alignment       =   1  'Right Justify
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
      Height          =   345
      Left            =   1770
      TabIndex        =   22
      Top             =   4140
      Width           =   2895
   End
   Begin VB.TextBox txt职工身份 
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
      Height          =   345
      Left            =   1770
      TabIndex        =   20
      Top             =   3690
      Width           =   2895
   End
   Begin VB.TextBox txt个人属地 
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
      Height          =   345
      Left            =   1770
      TabIndex        =   18
      Top             =   4140
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox txt卡状态 
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
      Height          =   345
      Left            =   1770
      TabIndex        =   16
      Top             =   3690
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.TextBox txt出生日期 
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
      Left            =   1770
      TabIndex        =   14
      Top             =   3210
      Width           =   2865
   End
   Begin VB.TextBox txt性别 
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
      Height          =   345
      Left            =   1770
      TabIndex        =   10
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox txt身份证号 
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
      Left            =   1770
      TabIndex        =   12
      Top             =   2730
      Width           =   2865
   End
   Begin VB.TextBox txt姓名 
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
      Height          =   345
      Left            =   1770
      TabIndex        =   8
      Top             =   1830
      Width           =   1995
   End
   Begin VB.TextBox txt卡号 
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
      Height          =   345
      Left            =   1770
      TabIndex        =   6
      Top             =   1380
      Width           =   2865
   End
   Begin VB.TextBox txt医保编号 
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
      Height          =   345
      Left            =   1770
      TabIndex        =   2
      Top             =   480
      Width           =   2865
   End
   Begin VB.Label lbl密码 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "密码(&P)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   810
      TabIndex        =   3
      Top             =   975
      Width           =   840
   End
   Begin VB.Label lbl医保中心 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "医保中心(&Z)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   330
      TabIndex        =   23
      Top             =   4635
      Width           =   1320
   End
   Begin VB.Label lbl帐户余额 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "帐户余额(&E)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   330
      TabIndex        =   21
      Top             =   4185
      Width           =   1320
   End
   Begin VB.Label lbl职工身份 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "职工身份(&F)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   330
      TabIndex        =   19
      Top             =   3735
      Width           =   1320
   End
   Begin VB.Label lbl个人属地 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "个人属地(&D)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   330
      TabIndex        =   17
      Top             =   4185
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Label lbl卡状态 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "卡状态(&T)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   570
      TabIndex        =   15
      Top             =   3735
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Label lbl出生日期 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "出生日期(&B)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   330
      TabIndex        =   13
      Top             =   3270
      Width           =   1320
   End
   Begin VB.Label lbl身份证号 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "身份证号(&I)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   330
      TabIndex        =   11
      Top             =   2790
      Width           =   1320
   End
   Begin VB.Label lbl性别 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "性别(&S)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   810
      TabIndex        =   9
      Top             =   2325
      Width           =   840
   End
   Begin VB.Label lbl姓名 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "姓名(&X)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   810
      TabIndex        =   7
      Top             =   1875
      Width           =   840
   End
   Begin VB.Label lbl卡号 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "卡号(&K)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   810
      TabIndex        =   5
      Top             =   1425
      Width           =   840
   End
   Begin VB.Label lbl医保编号 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "医保编号(&Y)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   330
      TabIndex        =   1
      Top             =   525
      Width           =   1320
   End
End
Attribute VB_Name = "frmIdentify自贡"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrReturn As String
Private mbytType As Byte
Private mlng病人ID As Long
Private mbln判断在院 As Boolean
Private mbln修改密码 As Boolean
Private mstrPass As String      '保存当前密码

Private Function IsValid() As Boolean
'功能：判断IC卡是否合法
    Dim rsTemp As New ADODB.Recordset
    Dim str有效期 As String
    Dim bln定点医疗 As Boolean
    
    If Me.txt姓名.Text = "" Then
        MsgBox "医保病人的身份还未确认，请先读卡！", vbInformation, gstrSysName
        Exit Function
    End If
    
    'G） 判断职工是否在住院：判断IC卡中InpatientFlag。（住院结算不进行此判断）
    If mbln判断在院 = True Then
        gstrSQL = "Select Nvl(当前状态,0) AS 状态 From 保险帐户 Where 险类=[1] And 医保号=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "判断当前病人是否在院", TYPE_四川自贡, CStr(txt医保编号.Text))
        If rsTemp.RecordCount <> 0 Then
            If rsTemp!状态 = 1 Then
                MsgBox "当前病人目前正在院接受治疗，不允许再次入院！", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    
    IsValid = True
End Function

Public Function GetPatient(ByVal bytType As Byte, Optional lng病人ID As Long, _
    Optional ByVal bln判断在院 As Boolean, Optional ByVal bln修改密码 As Boolean = False) As String
    
    mstrReturn = ""
    mbytType = bytType
    mlng病人ID = lng病人ID
    mbln判断在院 = bln判断在院
    mbln修改密码 = bln修改密码
    
    frmIdentify自贡.Show vbModal
    lng病人ID = mlng病人ID
    GetPatient = mstrReturn
End Function

Private Sub chk离休人员_Click()
    txt医保编号.Enabled = (chk离休人员.Value = 1)
    If txt医保编号.Enabled Then
        txt医保编号.SetFocus
    Else
        txt密码.SetFocus
    End If
End Sub

Private Sub cmd取消_Click()
    Unload Me
End Sub

Private Sub cmd确认_Click()
    '调用子过程清单
    'IsValid:对必要状态进行检查
    
    Dim rsTemp As New ADODB.Recordset
    Dim rsSelected As New ADODB.Recordset
    Dim rs病种 As New ADODB.Recordset
    Dim strIdentify As String, strAddition As String, strBirthday As String
    Dim lng病种 As Long, str病种 As String
    Dim int当前状态 As Integer
    Dim datToday As Date
    
    '产生病人信息串
    If Not IsValid() Then Exit Sub
    
    '取该病人的当前状态
    gstrSQL = "Select Nvl(当前状态,0) AS 当前状态 From 保险帐户 Where 医保号=[1] And 险类=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取该病人的当前状态", CStr(Trim(txt医保编号.Text)), TYPE_四川眉山)
    int当前状态 = 0
    If rsTemp.RecordCount <> 0 Then
        int当前状态 = rsTemp!当前状态
    End If
    
    '建立病人档案信息，传入格式：
    '0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);
    '8.中心代码;9.顺序号;10人员身份;11帐户余额;12当前状态;13病种ID;14在职(0,1);15退休证号;16年龄段;17灰度级
    '18帐户增加累计,19帐户支出累计,20进入统筹累计,21统筹报销累计,22住院次数累计,23就诊类型 (1、急诊门诊)
    strIdentify = TrimStr(txt医保编号.Text)                               '0卡号
    strIdentify = strIdentify & ";" & TrimStr(txt医保编号.Text)   '1医保号
    strIdentify = strIdentify & ";"                               '2密码
    strIdentify = strIdentify & ";" & TrimStr(txt姓名.Text)       '3姓名
    strIdentify = strIdentify & ";" & txt性别.Text    '4性别
    
    strBirthday = TrimStr(txt出生日期.Text)
    datToday = zlDatabase.Currentdate
    If strBirthday = "" Then
        strBirthday = Format(datToday, "yyyy-MM-dd")
    Else
        strBirthday = Mid(strBirthday, 1, 4) & "-" & Mid(strBirthday, 5, 2) & "-" & Mid(strBirthday, 7, 2)
    End If
    strIdentify = strIdentify & ";" & strBirthday              '5出生日期
    strIdentify = strIdentify & ";" & TrimStr(txt身份证号.Text)    '6身份证
    strIdentify = strIdentify & ";" & TrimStr(txt个人属地.Tag) & "(" & TrimStr(txt个人属地.Tag) & ")"   '7.单位名称(编码)
    
    '得到原住院病种
    If mbytType <> 1 Then
        gstrSQL = "Select Nvl(病种ID,0) 病种ID From 保险帐户 Where 险类=[1] And 医保号=[2]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "得到原住院病种", TYPE_四川自贡, CStr(TrimStr(txt医保编号.Text)))
        If Not rsTemp.EOF Then
            lng病种 = rsTemp!病种ID
        End If
    End If

    strAddition = ";" & Val(txt医保中心.Tag)                    '8.中心代码
    strAddition = strAddition & ";"                             '9.顺序号
    strAddition = strAddition & ";" & Mid(txt职工身份.Text, 1, InStr(1, txt职工身份.Text, "-") - 1)   '10人员身份
    strAddition = strAddition & ";" & Val(txt帐户余额.Text)   '11帐户余额
    strAddition = strAddition & ";" & int当前状态             '12当前状态
    strAddition = strAddition & ";" & IIf(lng病种 > 0, lng病种, "") '13病种ID

    strAddition = strAddition & ";" & Val(txt职工身份.Tag)
    strAddition = strAddition & ";"                             '15退休证号
    strAddition = strAddition & ";" & DateDiff("yyyy", CDate(strBirthday), datToday) '16年龄段
    strAddition = strAddition & ";"                             '17灰度级
    strAddition = strAddition & ";" & Val(txt帐户余额.Text)     '18帐户增加累计
    strAddition = strAddition & ";"                             '19帐户支出累计
    strAddition = strAddition & ";"                             '20进入统筹累计
    strAddition = strAddition & ";"                             '21统筹报销累计
    strAddition = strAddition & ";"                             '22住院次数累计
    strAddition = strAddition & ";"                             '23就诊类型 (1、急诊门诊)
    
    mlng病人ID = BuildPatiInfo(mbytType, strIdentify & strAddition, mlng病人ID, TYPE_四川自贡)
    '返回格式:中间插入病人ID
    mstrReturn = strIdentify & ";" & mlng病人ID & strAddition
    
    If mbytType = 1 Then
'        gstrSQL = "zl_病种信息_INSERT(" & TYPE_四川自贡 & "," & mlng病人ID & ",'" & str病种 & "')"
'        gcn自贡.Execute gstrSQL, , adCmdStoredProc
    End If
    
    Unload Me
End Sub

Private Sub cmd修改密码_Click()
    Dim strPass As String
    Dim StrInput As String, strOutput As String
    
    strPass = frm修改密码.ChangePassword("")
    If strPass = "" Then Exit Sub
    
    '调用修改密码函数
    StrInput = Me.txt密码.Text & "|" & strPass
    If Not 调用接口_自贡(业务类型_自贡.修改密码, StrInput, strOutput) Then Exit Sub
    
    '更新当前界面上的密码
    Me.txt密码.Text = strPass
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    cmd修改密码.Visible = mbln修改密码
End Sub

Private Sub txt密码_KeyDown(KeyCode As Integer, Shift As Integer)
    '调用子过程清单
    'Get离休人员:将离休人员的基本信息按读卡接口的返回格式组织起来
    
    Dim StrInput As String
    Dim strOutput As String
    Dim arrOutput
    Dim rsTemp As New ADODB.Recordset
    
    '读卡函数返回值常数
    Const cint医保编号 As Integer = 0
    Const cint卡号 As Integer = 1
    Const cint姓名 As Integer = 2
    Const cint性别 As Integer = 3
    Const cint身份证号 As Integer = 4
    Const cint出生日期 As Integer = 5
    Const cint卡状态 As Integer = 6
    Const cint个人属地 As Integer = 7
    Const cint单位编码 As Integer = 8
    Const cint职工身份 As Integer = 9
    Const cint帐户余额 As Integer = 10
    Const cint中心代码 As Integer = 11
    Const cint医保年度 As Integer = 12
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    
'    If chk离休人员.Value = 0 Then
        StrInput = Me.txt密码.Text
        If Not 调用接口_自贡(业务类型_自贡.读卡, StrInput, strOutput) Then Exit Sub
'    Else
'        strInput = Trim(txt医保编号.Text)
'        strOutPut = Get离休人员(strInput)
'        If strOutPut = "" Then
'            txt医保编号.SetFocus
'            Exit Sub
'        End If
'    End If
    '医保编号|卡号|姓名|性别|身份证号|出生日期|卡状态(标记)|个人属地信息|用人单位编码
    '|职工身份：0x-在职；1x-退休, 05和11为一次性缴费,7x二等乙级伤残军人|个人账户金额|中心代码|医保年度
    arrOutput = Split(strOutput, "|")
    
    '提取医保中心的名称与序号
    gstrSQL = "Select 序号,名称 From 保险中心目录 Where 编码=[1] And 险类=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取医保中心的名称与序号", CStr(arrOutput(cint中心代码)), TYPE_四川自贡)
    If rsTemp.RecordCount = 0 Then
        MsgBox "不存在该医保中心，中心代码为：" & arrOutput(cint中心代码), vbInformation, gstrSysName
        Exit Sub
    End If
    txt医保中心.Text = rsTemp!名称
    txt医保中心.Tag = Val(rsTemp!序号)
    Me.txt医保编号.Tag = arrOutput(cint中心代码)            '医保编号.tag保存中心代码
    
    Me.txt医保编号.Text = arrOutput(cint医保编号)
    Me.txt卡号.Text = arrOutput(cint卡号)
    Me.txt姓名.Text = arrOutput(cint姓名)
    Me.txt性别.Text = arrOutput(cint性别)
    Me.txt身份证号.Text = arrOutput(cint身份证号)
    Me.txt出生日期.Text = arrOutput(cint出生日期)
    Me.txt卡状态.Text = arrOutput(cint卡状态)
    Me.txt个人属地.Text = arrOutput(cint个人属地)
    Me.txt个人属地.Tag = arrOutput(cint单位编码)            '单位编码保存的txt个人属地的Tag中
    
    If arrOutput(cint职工身份) Like "5*" Then
        Me.txt职工身份.Text = arrOutput(cint职工身份) & "-" & "离休"
        Me.txt职工身份.Tag = 3
    ElseIf arrOutput(cint职工身份) Like "0*" Then
        Me.txt职工身份.Text = arrOutput(cint职工身份) & "-" & "在职"
        Me.txt职工身份.Tag = 1
    ElseIf arrOutput(cint职工身份) Like "7*" Then
        Me.txt职工身份.Text = arrOutput(cint职工身份) & "-" & "二等乙级伤残军人"
        Me.txt职工身份.Tag = 4
    Else
        Me.txt职工身份.Text = arrOutput(cint职工身份) & "-" & "退休"
        Me.txt职工身份.Tag = 2
    End If
    
    Me.txt帐户余额.Text = Format(Val(arrOutput(cint帐户余额)), "#####0.00")
End Sub

Private Sub txt医保编号_GotFocus()
    Call zlControl.TxtSelAll(txt医保编号)
End Sub
