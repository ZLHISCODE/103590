VERSION 5.00
Begin VB.Form frmIdentify吉林 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "病人身份验证"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6675
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   6675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraInfo 
      Caption         =   "病人信息"
      Height          =   4740
      Left            =   90
      TabIndex        =   35
      Top             =   115
      Width           =   6420
      Begin VB.CommandButton Cmd病种1 
         Caption         =   "…"
         Height          =   285
         Left            =   5895
         TabIndex        =   41
         Top             =   3960
         Width           =   255
      End
      Begin VB.TextBox txtInfo 
         Height          =   300
         Index           =   16
         Left            =   1290
         TabIndex        =   40
         Top             =   3960
         Width           =   4890
      End
      Begin VB.TextBox Txt疾病 
         Height          =   300
         Left            =   1290
         TabIndex        =   37
         Top             =   4335
         Width           =   4890
      End
      Begin VB.CommandButton cmd病种 
         Caption         =   "…"
         Height          =   285
         Left            =   5895
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   3600
         Width           =   255
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   15
         Left            =   4515
         TabIndex        =   31
         Top             =   3255
         Width           =   1650
      End
      Begin VB.ComboBox cboType 
         Height          =   300
         Left            =   1290
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   3255
         Width           =   1650
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   12
         Left            =   1290
         TabIndex        =   25
         Top             =   2820
         Width           =   1650
      End
      Begin VB.TextBox txtInfo 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Index           =   10
         Left            =   1290
         TabIndex        =   21
         Top             =   2390
         Width           =   1650
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   8
         Left            =   1290
         TabIndex        =   17
         Top             =   1960
         Width           =   1650
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   5
         Left            =   4515
         TabIndex        =   11
         Top             =   1100
         Width           =   1650
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   3
         Left            =   4515
         TabIndex        =   7
         Top             =   670
         Width           =   1650
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   4515
         TabIndex        =   3
         Top             =   240
         Width           =   1650
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   13
         Left            =   4515
         TabIndex        =   27
         Top             =   2820
         Width           =   1650
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   11
         Left            =   4515
         TabIndex        =   23
         Top             =   2390
         Width           =   1650
      End
      Begin VB.TextBox txtInfo 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Index           =   9
         Left            =   4515
         TabIndex        =   19
         Top             =   1960
         Width           =   1650
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   7
         Left            =   4515
         TabIndex        =   15
         Top             =   1530
         Width           =   1650
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   6
         Left            =   1290
         TabIndex        =   13
         Top             =   1530
         Width           =   1650
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   4
         Left            =   1290
         TabIndex        =   9
         Top             =   1100
         Width           =   1650
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   2
         Left            =   1290
         TabIndex        =   5
         Top             =   670
         Width           =   1650
      End
      Begin VB.TextBox txtInfo 
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   1290
         TabIndex        =   1
         Top             =   240
         Width           =   1650
      End
      Begin VB.TextBox txtInfo 
         Height          =   300
         Index           =   14
         Left            =   1290
         TabIndex        =   42
         Top             =   3600
         Width           =   4890
      End
      Begin VB.Label lblinfo 
         AutoSize        =   -1  'True
         Caption         =   "病种I"
         Height          =   180
         Index           =   15
         Left            =   810
         TabIndex        =   43
         Top             =   3660
         Width           =   450
      End
      Begin VB.Label lblinfo 
         AutoSize        =   -1  'True
         Caption         =   "病种II"
         Height          =   180
         Index           =   17
         Left            =   720
         TabIndex        =   39
         Top             =   4020
         Width           =   540
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "诊断情况"
         Height          =   180
         Index           =   16
         Left            =   540
         TabIndex        =   38
         Top             =   4395
         Width           =   720
      End
      Begin VB.Label lblinfo 
         AutoSize        =   -1  'True
         Caption         =   "支付序列号"
         Height          =   180
         Index           =   16
         Left            =   3600
         TabIndex        =   30
         Top             =   3330
         Width           =   900
      End
      Begin VB.Label lblinfo 
         AutoSize        =   -1  'True
         Caption         =   "交易类型"
         Height          =   180
         Index           =   7
         Left            =   540
         TabIndex        =   28
         Top             =   3330
         Width           =   720
      End
      Begin VB.Label lblinfo 
         AutoSize        =   -1  'True
         Caption         =   "是否重大疾病"
         Height          =   180
         Index           =   13
         Left            =   180
         TabIndex        =   24
         Top             =   2895
         Width           =   1080
      End
      Begin VB.Label lblinfo 
         AutoSize        =   -1  'True
         Caption         =   "住院次数"
         Height          =   180
         Index           =   11
         Left            =   540
         TabIndex        =   20
         Top             =   2460
         Width           =   720
      End
      Begin VB.Label lblinfo 
         AutoSize        =   -1  'True
         Caption         =   "人员身份"
         Height          =   180
         Index           =   9
         Left            =   540
         TabIndex        =   16
         Top             =   2040
         Width           =   720
      End
      Begin VB.Label lblinfo 
         AutoSize        =   -1  'True
         Caption         =   "姓名"
         Height          =   180
         Index           =   4
         Left            =   4080
         TabIndex        =   6
         Top             =   750
         Width           =   360
      End
      Begin VB.Label lblinfo 
         AutoSize        =   -1  'True
         Caption         =   "身份证号码"
         Height          =   180
         Index           =   3
         Left            =   3540
         TabIndex        =   10
         Top             =   1170
         Width           =   900
      End
      Begin VB.Label lblinfo 
         AutoSize        =   -1  'True
         Caption         =   "医保帐号"
         Height          =   180
         Index           =   1
         Left            =   3720
         TabIndex        =   2
         Top             =   315
         Width           =   720
      End
      Begin VB.Label lblinfo 
         AutoSize        =   -1  'True
         Caption         =   "住院标志"
         Height          =   180
         Index           =   14
         Left            =   3720
         TabIndex        =   26
         Top             =   2895
         Width           =   720
      End
      Begin VB.Label lblinfo 
         AutoSize        =   -1  'True
         Caption         =   "是否慢性病"
         Height          =   180
         Index           =   12
         Left            =   3540
         TabIndex        =   22
         Top             =   2460
         Width           =   900
      End
      Begin VB.Label lblinfo 
         AutoSize        =   -1  'True
         Caption         =   "帐户余额"
         Height          =   180
         Index           =   10
         Left            =   3720
         TabIndex        =   18
         Top             =   2040
         Width           =   720
      End
      Begin VB.Label lblinfo 
         AutoSize        =   -1  'True
         Caption         =   "出生日期"
         Height          =   180
         Index           =   8
         Left            =   3720
         TabIndex        =   14
         Top             =   1605
         Width           =   720
      End
      Begin VB.Label lblinfo 
         AutoSize        =   -1  'True
         Caption         =   "单位编码"
         Height          =   180
         Index           =   6
         Left            =   540
         TabIndex        =   12
         Top             =   1605
         Width           =   720
      End
      Begin VB.Label lblinfo 
         AutoSize        =   -1  'True
         Caption         =   "性别"
         Height          =   180
         Index           =   5
         Left            =   900
         TabIndex        =   8
         Top             =   1170
         Width           =   360
      End
      Begin VB.Label lblinfo 
         AutoSize        =   -1  'True
         Caption         =   "IC卡号"
         Height          =   180
         Index           =   2
         Left            =   720
         TabIndex        =   4
         Top             =   750
         Width           =   540
      End
      Begin VB.Label lblinfo 
         AutoSize        =   -1  'True
         Caption         =   "中心代码"
         Height          =   180
         Index           =   0
         Left            =   540
         TabIndex        =   0
         Top             =   315
         Width           =   720
      End
   End
   Begin VB.CommandButton cmdRead 
      Caption         =   "读卡(&R)"
      Height          =   400
      Left            =   210
      TabIndex        =   32
      Top             =   5055
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   400
      Left            =   4245
      TabIndex        =   33
      Top             =   5055
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   400
      Left            =   5400
      TabIndex        =   34
      Top             =   5055
      Width           =   1100
   End
End
Attribute VB_Name = "frmIdentify吉林"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbytType  As Byte   'bytType-识别类型，0-门诊收费，1-入院登记，2-不区分门诊与住院,3-挂号,4-结帐
Private mlng病人ID As Long
Private mstrReturn As String

Public Function GetPatient(bytType As Byte, lng病人ID As Long) As String
'参数：bytType-识别类型，0-门诊收费，1-入院登记，2-不区分门诊与住院,3-挂号,4-结帐
    mbytType = bytType
    mlng病人ID = lng病人ID
    mstrReturn = ""
    Me.Show vbModal
    GetPatient = mstrReturn
    lng病人ID = mlng病人ID
End Function

Private Sub cboType_Click()
    If cboType.ItemData(cboType.ListIndex) <> 2 And cboType.ItemData(cboType.ListIndex) <> 3 And mbytType = 0 Then
        cmd病种.Enabled = False
        txtInfo(14).Enabled = False
        
        '陈宏悦修改于20060512
        Cmd病种1.Enabled = False
        txtInfo(16).Enabled = False
    Else
        cmd病种.Enabled = True
        txtInfo(14).Enabled = True
    
        '陈宏悦修改于20060512
        Cmd病种1.Enabled = True
        txtInfo(16).Enabled = True
        
    End If
End Sub

Private Sub cboType_KeyDown(KeyCode As Integer, Shift As Integer)
    '刘兴宏:20040923加入
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Function IsValid() As Boolean
    '刘兴宏:20040923,加入验卡效正.
    Dim rsTemp As New ADODB.Recordset
    IsValid = False
    If txtInfo(0).Text = "" Then
        MsgBox "请读卡。", vbInformation, "读卡"
        Exit Function
    End If
    
    If mbytType = 0 And (cboType.ItemData(cboType.ListIndex) = 2 Or cboType.ItemData(cboType.ListIndex) = 3) And txtInfo(14).Tag = "" And txtInfo(16).Tag = "" Then
        MsgBox "请输入（或选择）正确的保险病种", vbInformation, "身份验证"
        Exit Function
    End If
    If mbytType = 0 And Trim(Txt疾病.Tag) = "" Then
        ShowMsgbox "请输入诊断情况"
        Exit Function
    End If
    If mbytType <> 2 Then
        If mbytType = 4 Then
            '不检查当前着态
        Else
            '检查病人状态
            gstrSQL = "select nvl(当前状态,0) as 状态 from 保险帐户 where 险类=[1] and 医保号=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, TYPE_吉林, CStr(txtInfo(1).Text))
            If rsTemp.RecordCount > 0 Then
                If rsTemp("状态") > 0 Then
                    MsgBox "该病人已经在院，不能通过身份验证。", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
        If mbytType = 0 Or mbytType = 3 Then
            '设置
            
        End If
    Else
        '不区分门诊和住院的，只是刷卡显示一下内容而已，不保存
        Unload Me
        Exit Function
    End If
        
    IsValid = True
End Function
Private Sub cmdOK_Click()
    Dim strTmp As String
    Dim strAddition As String, strIdentify As String
    If IsValid = False Then Exit Sub
    
    If cboType.ListIndex >= 0 Then
        g病人身份_吉林.门诊类型 = cboType.ItemData(cboType.ListIndex)
    End If
    
    With g病人身份_吉林
        
        If Val(txtInfo(14).Tag) <> 0 Or Val(txtInfo(16).Tag) = 0 Then
        
            .病种代码 = cmd病种.Tag
            
        End If
        
        If Val(txtInfo(14).Tag) = 0 Or Val(txtInfo(16).Tag) <> 0 Then
        
            .病种代码 = Cmd病种1.Tag
            
        End If
        
        If Val(txtInfo(14).Tag) <> 0 Or Val(txtInfo(16).Tag) <> 0 Then
        
        '陈宏悦于20060228修改,由于银海医保调整门诊慢性病
        '多个单病种分割方式,以逗号做分隔符
        
            .病种代码 = cmd病种.Tag & "," & Cmd病种1.Tag
            
        End If
        
        If Txt疾病.Tag <> "" Then
        
        '陈宏悦于20060228修改,由于银海医保调整门诊慢性病
        '多个单病种分割方式,以逗号做分隔符
        
            .诊断编码 = Split(Txt疾病.Tag, "|||")(0)
            .诊断名称 = Split(Txt疾病.Tag, "|||")(1)
        Else
            If mbytType = 1 Then
                If Trim(Txt疾病.Text) = "" Then
                    ShowMsgbox "请输入诊断情况"
                    If Txt疾病.Enabled Then Txt疾病.SetFocus
                    Exit Sub
                End If
                .诊断名称 = Txt疾病
            End If
        End If
        
        strAddition = "": strIdentify = ""
        strIdentify = .卡号                               '0卡号
        strIdentify = strIdentify & ";" & .医保号            '1医保号
        strIdentify = strIdentify & ";" & .门诊类型            '2密码  交易类型
        strIdentify = strIdentify & ";" & .姓名               '3姓名
        strIdentify = strIdentify & ";" & .性别                 '4性别
        strIdentify = strIdentify & ";" & .出生日期                    '5出生日期
        
        strIdentify = strIdentify & ";" & .身份证号                     '6身份证
        strIdentify = strIdentify & ";" & "(" & .单位编码 & ")"                 '7.单位名称(编码)
        strAddition = ";0"                                          '8.中心代码
        strAddition = strAddition & ";" & .支付序列                              '9 顺序号
        
        strAddition = strAddition & ";" & .个人身份                                           '10 人员身份
        strAddition = strAddition & ";" & .帐户余额                        '11 余额
        
        strAddition = strAddition & ";0"                                        '12 当前状态
        strAddition = strAddition & ";" & txtInfo(14).Tag                        '13 病种ID
        strAddition = strAddition & ";" & IIf(txtInfo(8).Text = "在职", "1", "2")   '14在职(1,2,3)
        
        strTmp = .公务员标志 & "|" & .补充保险 & "|" & .大病医保 & "|" & .隶属关系 & "|" & .照顾级别 & "|" & .职工属地 & "|" & .是否慢性病 & " |" & .重大疾病
        strAddition = strAddition & ";" & strTmp                                '15 退休证号
        strAddition = strAddition & ";"                                         '16 年龄段
        strAddition = strAddition & ";"                                         '17 灰度级
        strAddition = strAddition & ";" & .帐户余额                             '18 帐户增加累计
        strAddition = strAddition & ";" & "0"                                         '19 帐户支出累计
        
        strAddition = strAddition & ";" & .进入统筹累计                        '20 进入统筹累计
        strAddition = strAddition & ";" & .统筹支付累计                        '21 统筹报销累计
        strAddition = strAddition & ";" & .住院次数                            '22 住院次数累计
        strAddition = strAddition & ";"                                      '23 就诊类别
        strAddition = strAddition & ";"                                      '24 本次起付线
        strAddition = strAddition & .起付线金额累计 & ";"                         '25 起付线累计
        strAddition = strAddition & ";" & .起付段医疗费累计                  '26 基本统筹限额
        
    End With
    'Me.Hide
    '刘兴宏:20040923加入
    mlng病人ID = BuildPatiInfo(0, strIdentify & strAddition, mlng病人ID, TYPE_吉林)
    
    '更新诊断情况
    gstrSQL = "ZL_保险帐户_更新信息(" & mlng病人ID & "," & TYPE_吉林 & ",'入院诊断','''" & g病人身份_吉林.诊断名称 & "''')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "更新诊断情况")
     
    If mlng病人ID > 0 Then
        mstrReturn = strIdentify & ";" & mlng病人ID & strAddition
    End If
    Unload Me
End Sub

Private Sub cmdRead_Click()
    If 身份鉴别_吉林(mbytType) = False Then Exit Sub
    
    '显示读卡所获信息
    With g病人身份_吉林
        txtInfo(0).Text = .中心      '中心
        txtInfo(2).Text = .卡号      '卡号
        txtInfo(5).Text = .身份证号      '身份证号
        txtInfo(3).Text = .姓名    '姓名
        txtInfo(4).Text = .性别
        txtInfo(7).Text = .出生日期
        txtInfo(1).Text = .医保号      '医保号
        txtInfo(6).Text = .单位编码
        txtInfo(8).Text = IIf(Val(.个人身份) = "0", "在职", "退休")
        
        txtInfo(9).Text = .帐户余额
        txtInfo(10).Text = .住院次数
        txtInfo(11).Text = IIf(.是否慢性病 = "0", "不是", "是")
        txtInfo(12).Text = IIf(.重大疾病 = "0", "不是", "是")
        txtInfo(13).Text = IIf(.住院标志 = "0", "不住院", "住院")
        '刘兴宏:20040923加入
        txtInfo(15).Text = .支付序列
    End With
End Sub

 Private Sub SetCtlBackColor()
    '刘兴宏:20040923加入
    Dim i As Long
    For i = 0 To txtInfo.UBound
        If i <> 14 Or i <> 16 Then
        Else
            txtInfo(i).BackColor = &H8000000F
        End If
    Next
 End Sub

Private Sub cmd病种_Click()
    Dim rsTemp As New ADODB.Recordset
    
    If mbytType = 0 Or mbytType = 3 Then
        '只选慢病和特种病
        gstrSQL = " Select A.ID,编码,A.名称,A.简码,decode(A.类别,'1','慢性病','2','特种病','普通病') as 类别 " & _
                "   From 保险病种 A " & _
                "   where nvl(a.类别,0)<>'0' and A.险类=" & TYPE_吉林 & _
                "   order by 编码"
    ElseIf mbytType = 1 Or mbytType = 4 Then
        gstrSQL = " Select A.ID,编码,A.名称,A.简码,decode(A.类别,'1','慢性病','2','特种病','普通病') as 类别 " & _
                "   From 保险病种 A " & _
                "   where nvl(a.类别,0)='0' and A.险类=" & TYPE_吉林 & _
                "   order by 编码"
    Else
        gstrSQL = " Select A.ID,编码,A.名称,A.简码,decode(A.类别,'1','慢性病','2','特种病','普通病') as 类别 " & _
                "   From 保险病种 A " & _
                "   where A.险类=" & TYPE_吉林 & _
                "   order by 编码"
    End If
    
    Set rsTemp = frmPubSel.ShowSelect(Me, gstrSQL, 0, "医保病种", , txtInfo(14).Text)
    If rsTemp.State = 0 Then Exit Sub
    
    If Not rsTemp Is Nothing Then
        txtInfo(14).Text = "(" & rsTemp!编码 & ")" & rsTemp!名称
        txtInfo(14).Tag = rsTemp("ID")
        cmd病种.Tag = Nvl(rsTemp!编码)
        zlControl.TxtSelAll txtInfo(14)
    End If
    txtInfo(14).SetFocus
End Sub

Private Sub Cmd病种1_Click()
   
'陈宏悦于20060228修改

   Dim rsTemp As New ADODB.Recordset
    
    If mbytType = 0 Or mbytType = 3 Then
        '只选慢病和特种病
        gstrSQL = " Select A.ID,编码,A.名称,A.简码,decode(A.类别,'1','慢性病','2','特种病','普通病') as 类别 " & _
                "   From 保险病种 A " & _
                "   where nvl(a.类别,0)<>'0' and A.险类=" & TYPE_吉林 & _
                "   order by 编码"
    ElseIf mbytType = 1 Or mbytType = 4 Then
        gstrSQL = " Select A.ID,编码,A.名称,A.简码,decode(A.类别,'1','慢性病','2','特种病','普通病') as 类别 " & _
                "   From 保险病种 A " & _
                "   where nvl(a.类别,0)='0' and A.险类=" & TYPE_吉林 & _
                "   order by 编码"
    Else
        gstrSQL = " Select A.ID,编码,A.名称,A.简码,decode(A.类别,'1','慢性病','2','特种病','普通病') as 类别 " & _
                "   From 保险病种 A " & _
                "   where A.险类=" & TYPE_吉林 & _
                "   order by 编码"
    End If
    
    Set rsTemp = frmPubSel.ShowSelect(Me, gstrSQL, 0, "医保病种", , txtInfo(16).Text)
    If rsTemp.State = 0 Then Exit Sub
    
    If Not rsTemp Is Nothing Then
        txtInfo(16).Text = "(" & rsTemp!编码 & ")" & rsTemp!名称
        txtInfo(16).Tag = rsTemp("ID")
        Cmd病种1.Tag = Nvl(rsTemp!编码)
        zlControl.TxtSelAll txtInfo(16)
    End If
    txtInfo(16).SetFocus
End Sub

Private Sub Form_Load()
    '刘兴宏:没有必要设置相关的
    cboType.Clear
    '陈宏悦于20060512修改
    g病人身份_吉林.病种代码 = ""
    
    If mbytType = 0 Or mbytType = 3 Then
        cboType.Enabled = True
        txtInfo(14).Enabled = True
        Txt疾病.Enabled = True
        
        '陈宏悦于20060228修改
        '原因：主要是由于银海公司对于慢性病可能出现两种
        txtInfo(16).Enabled = True
                
        cboType.AddItem "1-普通"
        cboType.ItemData(cboType.NewIndex) = 1
        cboType.AddItem "2-慢性病"
        cboType.ItemData(cboType.NewIndex) = 2
        cboType.AddItem "3-重大疾病"
        cboType.ItemData(cboType.NewIndex) = 3
        cboType.AddItem "4-照顾对象"
        cboType.ItemData(cboType.NewIndex) = 4
        cboType.AddItem "5-特种人"
        cboType.ItemData(cboType.NewIndex) = 5
        cboType.AddItem "6-计划生育"
        cboType.ItemData(cboType.NewIndex) = 6
        cboType.AddItem "7-工伤"
        cboType.ItemData(cboType.NewIndex) = 7
        cboType.ListIndex = 0
    Else
        If mbytType = 1 Then
            Txt疾病.Enabled = True
        Else
            Txt疾病.Enabled = False
        End If
        
        '陈宏悦于20060228修改
        txtInfo(16).Enabled = False
        Cmd病种1.Enabled = False
        
        cboType.AddItem "1-普通"
        cboType.ItemData(cboType.NewIndex) = 1
        cboType.AddItem "2.计划生育住院结算"
        cboType.ItemData(cboType.NewIndex) = 6
        cboType.ListIndex = 0
    End If
    
    Call SetCtlBackColor
End Sub

Private Sub txtInfo_Change(Index As Integer)
    txtInfo(Index).Tag = ""
        
End Sub

Private Sub txtInfo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    '刘兴宏:20040923加入
    Dim strLike As String, StrInput As String
    Dim blnCancel As Boolean
    Dim rsTmp As ADODB.Recordset
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    '陈宏悦于20060228修改
    If Index = 14 Or Index = 16 Then
        '选择病种
        If txtInfo(Index).Text = "" Then
            Call zlCommFun.PressKey(vbKeyTab) '允许不输入
        Else
            strLike = IIf(GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", 0) = 0, "%", "")
            StrInput = strLike & UCase(txtInfo(Index).Text)
            
            If mbytType = 0 Or mbytType = 3 Then
                '只选慢病和特种病
                gstrSQL = " Select A.ID,编码,A.名称,A.简码,decode(A.类别,'1','慢性病','2','特种病','普通病') as 类别 " & _
                        "   From 保险病种 A " & _
                        "   where nvl(a.类别,0)<>'0' and A.险类=" & TYPE_吉林 & _
                        "        and ( 编码 like '" & StrInput & "%' or 名称 like '" & StrInput & "%' or 简码 like '" & StrInput & "%' )" & _
                        "        And Rownum<=200" & _
                        "   order by 编码"
            ElseIf mbytType = 1 Or mbytType = 4 Then
                gstrSQL = " Select A.ID,编码,A.名称,A.简码,decode(A.类别,'1','慢性病','2','特种病','普通病') as 类别 " & _
                        "   From 保险病种 A " & _
                        "   where nvl(a.类别,0)='0' and A.险类=" & TYPE_吉林 & _
                        "        and ( 编码 like '" & StrInput & "%' or 名称 like '" & StrInput & "%' or 简码 like '" & StrInput & "%' )" & _
                        "        And Rownum<=200" & _
                        "   order by 编码"
            Else
                gstrSQL = " Select A.ID,编码,A.名称,A.简码,decode(A.类别,'1','慢性病','2','特种病','普通病') as 类别 " & _
                        "   From 保险病种 A " & _
                        "   where A.险类=" & TYPE_吉林 & _
                        "        and ( 编码 like '" & StrInput & "%' or 名称 like '" & StrInput & "%' or 简码 like '" & StrInput & "%' )" & _
                        "        And Rownum<=200" & _
                        "   order by 编码"
            End If
            Set rsTmp = zlDatabase.ShowSelect(Me, gstrSQL, 0, "病种选择", , , , , , True, _
                txtInfo(Index).Left + Me.Left, _
                txtInfo(Index).Top + Me.Top, txtInfo(Index).Height, blnCancel, , True)
                
            If Not rsTmp Is Nothing Then
                txtInfo(Index).Text = "(" & rsTmp!编码 & ")" & rsTmp!名称
                txtInfo(Index).Tag = rsTmp("ID")
                
                '陈宏悦于20060228修改
                If Index = 14 Then
                    cmd病种.Tag = Nvl(rsTmp!编码)
                ElseIf Index = 16 Then
                    Cmd病种1.Tag = Nvl(rsTmp!编码)
                End If
            Else
                If Not blnCancel Then
                    MsgBox "没有找到匹配的病种编码。", vbInformation, gstrSysName
                End If
                txtInfo(Index).SetFocus
            End If
        End If
    End If
    zlCommFun.PressKey vbKeyTab
End Sub


Private Sub Txt疾病_Change()
    Txt疾病.Tag = ""
End Sub

Private Sub Txt疾病_GotFocus()
    zlControl.TxtSelAll Txt疾病
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub Txt疾病_KeyPress(KeyAscii As Integer)
  Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, blnCancel As Boolean
    Dim strLike As String, str性别 As String
    Dim StrInput As String
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Txt疾病.Text = "" Then
            Call zlCommFun.PressKey(vbKeyTab) '允许不输入
        Else
            strLike = IIf(GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", 0) = 0, "%", "")
            StrInput = UCase(Txt疾病.Text)
            str性别 = g病人身份_吉林.性别
            If str性别 = "男" Then
                str性别 = " And (A.性别限制='男' Or A.性别限制 is NULL)"
            ElseIf str性别 = "女" Then
                str性别 = " And (A.性别限制='女' Or A.性别限制 is NULL)"
            End If
            
            strSQL = "Select A.ID,A.编码,A.附码,A.名称,A.简码,A.说明,A.性别限制,B.类别" & _
                " From 疾病编码目录 A,疾病编码类别 B" & _
                " Where A.类别=B.编码 And A.类别 Not IN('B','Z')" & _
                " And (A.编码 Like '" & StrInput & "%'" & _
                " Or Upper(A.名称) Like '" & strLike & StrInput & "%'" & _
                " Or Upper(A.简码) Like '" & strLike & StrInput & "%'" & _
                " Or Upper(A.附码) Like '" & strLike & StrInput & "%')" & _
                " And Rownum<=100" & str性别 & _
                " Order by A.类别,A.编码"
                
            Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, "疾病编码Input", , , , , , True, _
                Txt疾病.Left + Me.Left, _
                Txt疾病.Top + Me.Top, Txt疾病.Height, blnCancel, , True)
            If Not rsTmp Is Nothing Then
                Txt疾病.Text = "(" & rsTmp!编码 & ")" & rsTmp!名称
                Txt疾病.Tag = rsTmp!编码 & "|||" & rsTmp!名称
                If cmdOK.Enabled Then
                    cmdOK.SetFocus
                Else
                    Call zlCommFun.PressKey(vbKeyTab)
                End If
            Else
                If mbytType <> 1 Then
                    If Not blnCancel Then
                        MsgBox "没有找到匹配的疾病编码。", vbInformation, gstrSysName
                    End If
                    Call Txt疾病_GotFocus
                    Txt疾病.SetFocus
                Else
                        If cmdOK.Enabled Then
                            cmdOK.SetFocus
                        Else
                            Call zlCommFun.PressKey(vbKeyTab)
                        End If
                End If
            End If
        End If
    Else
        zlControl.TxtCheckKeyPress Txt疾病, KeyAscii, m文本式
    End If
End Sub


