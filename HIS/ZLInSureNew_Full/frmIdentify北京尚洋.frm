VERSION 5.00
Begin VB.Form frmIdentify北京尚洋 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "身份验证"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3300
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2265
   ScaleWidth      =   3300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame1 
      Caption         =   "病人信息"
      Height          =   1560
      Left            =   75
      TabIndex        =   5
      Top             =   75
      Width           =   3075
      Begin VB.TextBox txtNO 
         Height          =   300
         Left            =   990
         TabIndex        =   0
         Top             =   270
         Width           =   1890
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Left            =   990
         TabIndex        =   1
         Text            =   "医保病人"
         Top             =   660
         Width           =   1890
      End
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         Height          =   300
         Left            =   990
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1065
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "病人姓名"
         Height          =   180
         Index           =   1
         Left            =   180
         TabIndex        =   8
         Top             =   735
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "个人编号"
         Height          =   180
         Index           =   2
         Left            =   180
         TabIndex        =   7
         Top             =   345
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "性别"
         Height          =   180
         Index           =   3
         Left            =   540
         TabIndex        =   6
         Top             =   1140
         Width           =   360
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   400
      Left            =   945
      TabIndex        =   3
      Top             =   1755
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   400
      Left            =   2055
      TabIndex        =   4
      Top             =   1755
      Width           =   1100
   End
End
Attribute VB_Name = "frmIdentify北京尚洋"
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
    If Trim(txtNO.Text) = "" Then
        MsgBox "请输入医保病人的个人编号！", vbInformation, gstrSysName
        txtNO.SetFocus
        Exit Sub
    End If
'    If Trim(Text1.Text) = "" Then
'        If Text1.Enabled = True Then
'            MsgBox "请输入医保病人姓名", vbInformation, gstrSysName
'            Text1.SetFocus
'            Exit Sub
'        End If
'    End If
    WriteInfo Format(Date, "yyyy-MM-dd") & " " & Format(Time, "HH:MI:SS") & "开始生成病人身份信息"
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
    strIdentify = strIdentify & txtNO.Tag & ";"                 '5出生日期
    strIdentify = strIdentify & ";"                             '6身份证
    strIdentify = strIdentify & Text1.Tag & ";"                 '7.单位名称(编码)
    strAddition = "0;"                                          '8.中心代码
    strAddition = strAddition & ";"                             '9.顺序号
    strAddition = strAddition & Combo1.Tag & ";"                '10人员身份
    '因为不能取得个帐余额，所以将个帐余额赋予足够大的值，具体支付金额由医保确定
    strAddition = strAddition & "1000000;"                      '11帐户余额
    strAddition = strAddition & "0;"                            '12当前状态
    strAddition = strAddition & ";"                             '13病种ID
    strAddition = strAddition & "1;"                            '14在职(1,2,3)
    strAddition = strAddition & ";"                             '15退休证号
    strAddition = strAddition & ";"                             '16年龄段
    strAddition = strAddition & ";"                             '17灰度级
    strAddition = strAddition & "1000000;"                      '18帐户增加累计
    strAddition = strAddition & ";"                             '19帐户支出累计
    strAddition = strAddition & "0;"                            '20进入统筹累计
    strAddition = strAddition & "0;"                            '21统筹报销累计
    strAddition = strAddition & "0;"                            '22住院次数累计
    strAddition = strAddition & ";"                             '23就诊类型
    WriteInfo Format(Date, "yyyy-MM-dd") & " " & Format(Time, "HH:MI:SS") & "开始建立病人档案，数据： strIdentify & strAddition"
    mlng病人ID = BuildPatiInfo(0, strIdentify & strAddition, mlng病人ID, TYPE_北京尚洋)
    WriteInfo Format(Date, "yyyy-MM-dd") & " " & Format(Time, "HH:MI:SS") & "完成档案建立"
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
    If KeyAscii = 13 Then cmdOK.SetFocus
End Sub

Private Sub Form_Load()
    Combo1.AddItem "男"
    Combo1.AddItem "女"
    Combo1.ListIndex = 0
End Sub

Private Sub Text1_GotFocus()
    Call zlControl.TxtSelAll(Text1)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Combo1.SetFocus
End Sub

Private Sub txtNO_GotFocus()
    Call zlControl.TxtSelAll(txtNO)
End Sub

Private Sub txtNO_KeyPress(KeyAscii As Integer)
    Dim StrInput As String, strSQL As String, rsTmp As New ADODB.Recordset
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If txtNO <> "" Then
        txtNO.Text = UCase(txtNO.Text)
        If mbytType = 1 Then
            Set rsTmp = gcn尚洋.Execute("Select * From SICK_VISIT_INFO Where PERSONAL_NUMBER='" & txtNO.Text & "' And HOSPITAL_NUMBER='" & gstr医院编码 & "'")
        Else
            Set rsTmp = gcn尚洋.Execute("Select * From MED_RECEIPT_RECORD_MASTER Where PERSONAL_NUMBER='" & txtNO.Text & "' And HOSPITAL_NUMBER='" & gstr医院编码 & "'")
        End If
        If rsTmp.EOF Then
            Text1.Enabled = True
            Combo1.Enabled = True
            txtNO.Tag = ""
            Text1.Tag = ""
            Combo1.Tag = ""
            Text1.SetFocus
        Else
            Text1.Enabled = False
            Combo1.Enabled = False
            Text1.Text = rsTmp!Name                                    '姓名
            Text1.Tag = rsTmp!UNIT_NUMBER                              '单位编码
            Combo1.ListIndex = IIf(Trim(Nvl(rsTmp!Sex, "0")) = "1", 0, 1)
            If mbytType = 1 Then
                Call DebugTool("出生日期:" & rsTmp!BIRTHMONTH)
                txtNO.Tag = Mid(rsTmp!BIRTHMONTH, 1, 4) & "-" & Mid(rsTmp!BIRTHMONTH, 5, 2) & "-01"     '出生日期
            Else
                txtNO.Tag = Format(rsTmp!BIRTH_DATE, "yyyy-MM-dd")         '出生日期
            End If
            Combo1.Tag = rsTmp!PERSONAL_TYPE                           '人员类别
            cmdOK.SetFocus
        End If
    End If
End Sub

