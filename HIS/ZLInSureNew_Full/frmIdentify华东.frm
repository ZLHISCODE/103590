VERSION 5.00
Begin VB.Form frmIdentify华东 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "身份验证"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3405
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   3405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   400
      Left            =   2140
      TabIndex        =   5
      Top             =   1785
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   400
      Left            =   1035
      TabIndex        =   4
      Top             =   1785
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Caption         =   "病人信息"
      Height          =   1560
      Left            =   165
      TabIndex        =   0
      Top             =   105
      Width           =   3075
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         Height          =   300
         Left            =   2085
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1065
         Width           =   795
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   300
         Left            =   990
         MaxLength       =   3
         TabIndex        =   7
         Top             =   1065
         Width           =   525
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Left            =   990
         TabIndex        =   3
         Top             =   660
         Width           =   1890
      End
      Begin VB.TextBox txtNO 
         Height          =   300
         Left            =   990
         TabIndex        =   1
         Top             =   270
         Width           =   1890
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "性别"
         Height          =   180
         Index           =   3
         Left            =   1635
         TabIndex        =   9
         Top             =   1140
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "年龄"
         Height          =   180
         Index           =   0
         Left            =   540
         TabIndex        =   8
         Top             =   1140
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "IC卡号"
         Height          =   180
         Index           =   2
         Left            =   360
         TabIndex        =   6
         Top             =   345
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "病人姓名"
         Height          =   180
         Index           =   1
         Left            =   180
         TabIndex        =   2
         Top             =   735
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmIdentify华东"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng病人ID As Long, mstrIdentify As String, mbytType As Byte

Public Function Identify(ByVal bytType As Byte, Optional lng病人ID As Long) As String
    mbytType = bytType
    mlng病人ID = lng病人ID
    Me.Show 1
    Identify = mstrIdentify
End Function

Private Sub cmdCancel_Click()
    mstrIdentify = ""
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    Dim strSQL As String
    Dim strIdentify As String
    Dim strAddition As String
    Dim lngSequence As String
    On Error GoTo errHandle
    
    '判断是否输入医保病人姓名
    If Trim(Text1.Text) = "" Then
        If Text1.Enabled = True Then
            MsgBox "请输入医保病人姓名", vbInformation, gstrSysName
            Text1.SetFocus
        ElseIf Trim(txtNO.Text) = "" Then
            MsgBox "请输入医保病人卡号", vbInformation, gstrSysName
            txtNO.SetFocus
        End If
        Exit Sub
    End If
    If Not IsNumeric(Trim(Text2.Text)) And Trim(Text2.Text) <> "" Then
        MsgBox "年龄输入错误", vbInformation, gstrSysName
        Text2.SetFocus
        Exit Sub
    End If
    lngSequence = UCase(txtNO.Text)
'      strInfo='0卡号;1医保号;2密码;3姓名;4性别;5出生日期;6身份证;7单位名称(编码);
'      8中心;9.顺序号;10人员身份;11帐户余额;12当前状态;13病种ID;14在职(1,2,3);15退休证号;16年龄段;17灰度级
'      18帐户增加累计;19帐户支出累计;20进入统筹累计;21统筹报销累计;22住院次数累计;23就诊类别
'      24本次起付线;25起付线累计;26基本统筹限额
    
    strIdentify = lngSequence & ";"                             '0卡号
    strIdentify = strIdentify & lngSequence & ";"               '1医保号（个人编号）
    strIdentify = strIdentify & ";"                             '2密码
    strIdentify = strIdentify & Trim(Text1.Text) & ";"          '3姓名
    strIdentify = strIdentify & Combo1.Text & ";"               '4性别
    strIdentify = strIdentify & ";"                             '5出生日期
    strIdentify = strIdentify & ";"                             '6身份证
    strIdentify = strIdentify & ";"                             '7.单位名称(编码)
    strAddition = "0;"                                          '8.中心代码
    strAddition = strAddition & ";"                             '9.顺序号
    strAddition = strAddition & ";"                             '10人员身份
    '因为不能取得个帐余额，所以将个帐余额赋予足够大的值，具体支付金额由医保确定
    strAddition = strAddition & "1000000;"                      '11帐户余额
    strAddition = strAddition & "0;"                            '12当前状态
    strAddition = strAddition & ";"                             '13病种ID
    strAddition = strAddition & "1;"                            '14在职(1,2,3)
    strAddition = strAddition & ";"                             '15退休证号
    strAddition = strAddition & Text2.Text & ";"                '16年龄段
    strAddition = strAddition & ";"                             '17灰度级
    strAddition = strAddition & "1000000;"                      '18帐户增加累计
    strAddition = strAddition & ";"                             '19帐户支出累计
    strAddition = strAddition & "0;"                            '20进入统筹累计
    strAddition = strAddition & "0;"                            '21统筹报销累计
    strAddition = strAddition & "0;"                            '22住院次数累计
    strAddition = strAddition & ";"                             '23就诊类型
    
    mlng病人ID = BuildPatiInfo(0, strIdentify & strAddition, mlng病人ID, TYPE_华东)
    '返回格式:中间插入病人ID
    If mlng病人ID > 0 Then
        mstrIdentify = strIdentify & mlng病人ID & ";" & strAddition
    End If
    Me.Hide
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        On Error Resume Next
        cmdOK.Enabled = True
        cmdOK.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Combo1.AddItem "男"
    Combo1.AddItem "女"
    Combo1.ListIndex = 0
End Sub

Private Sub Text1_GotFocus()
    On Error Resume Next
    Call zlControl.TxtSelAll(Text1)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        On Error Resume Next
        Text2.Enabled = True
        Text2.SetFocus
    End If
End Sub

Private Sub Text2_GotFocus()
    On Error Resume Next
    Call zlControl.TxtSelAll(Text2)
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        On Error Resume Next
        Combo1.Enabled = True
        Combo1.SetFocus
    End If
End Sub

Private Sub txtNO_GotFocus()
    On Error Resume Next
    Call zlControl.TxtSelAll(txtNO)
End Sub

Private Sub txtNO_KeyPress(KeyAscii As Integer)
    Dim StrInput As String, strSQL As String, rsTemp As New ADODB.Recordset
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If txtNO <> "" Then
        txtNO.Text = UCase(txtNO.Text)
        gstrSQL = "Select * From 保险帐户 Where 卡号=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CStr(txtNO.Text))
        If rsTemp.EOF Then
            Text1.Enabled = True
            Text2.Enabled = True
            Combo1.Enabled = True
            Text1.SetFocus
        Else
            gstrSQL = "Select * From 病人信息 Where 病人id=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CLng(rsTemp!病人ID))
            If rsTemp.EOF Then
                Text1.Enabled = True
                Text2.Enabled = True
                Combo1.Enabled = True
                Text1.SetFocus
            Else
                Text1.Enabled = False
                Text2.Enabled = False
                Combo1.Enabled = False
                Text1.Text = rsTemp!姓名
                Text2.Text = Nvl(rsTemp!年龄)
                Combo1.ListIndex = IIf(Nvl(rsTemp!性别, "男") = "男", 0, 1)
                cmdOK.SetFocus
            End If
        End If
    End If
End Sub
