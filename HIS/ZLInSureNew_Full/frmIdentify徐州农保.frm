VERSION 5.00
Begin VB.Form frmIdentify徐州农保 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "病人身份验证"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3810
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   3810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   400
      Left            =   2463
      TabIndex        =   6
      Top             =   2535
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   400
      Left            =   1365
      TabIndex        =   5
      Top             =   2535
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   75
      Left            =   -120
      TabIndex        =   12
      Top             =   2295
      Width           =   3990
   End
   Begin VB.TextBox txtEdit 
      Enabled         =   0   'False
      Height          =   300
      Index           =   4
      Left            =   1193
      TabIndex        =   4
      Top             =   1860
      Width           =   2370
   End
   Begin VB.TextBox txtEdit 
      Enabled         =   0   'False
      Height          =   300
      Index           =   3
      Left            =   1193
      TabIndex        =   3
      Top             =   1425
      Width           =   2370
   End
   Begin VB.TextBox txtEdit 
      Enabled         =   0   'False
      Height          =   300
      Index           =   2
      Left            =   1193
      TabIndex        =   2
      Top             =   990
      Width           =   2370
   End
   Begin VB.TextBox txtEdit 
      Enabled         =   0   'False
      Height          =   300
      Index           =   1
      Left            =   1193
      TabIndex        =   1
      Top             =   585
      Width           =   2370
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   0
      Left            =   1193
      TabIndex        =   0
      Top             =   195
      Width           =   2370
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "年龄"
      Height          =   180
      Index           =   4
      Left            =   788
      TabIndex        =   11
      Top             =   1935
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "性别"
      Height          =   180
      Index           =   3
      Left            =   788
      TabIndex        =   10
      Top             =   1500
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "病人姓名"
      Height          =   180
      Index           =   2
      Left            =   428
      TabIndex        =   9
      Top             =   1065
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "病人医保号"
      Height          =   180
      Index           =   1
      Left            =   255
      TabIndex        =   8
      Top             =   660
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "医保住院号"
      Height          =   180
      Index           =   0
      Left            =   248
      TabIndex        =   7
      Top             =   270
      Width           =   900
   End
End
Attribute VB_Name = "frmIdentify徐州农保"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mbytType As Byte, mstrPatient As String, mstrOther As String, mint住院次数 As Integer
Private strTransNO As String, cur支出累计 As Currency, cur增加累计 As Currency, strPara As String, _
    strReturn As String, blnReadCard As Boolean
 
Public Function GetPatient(bytType As Byte) As String
'参数：bytType-识别类型，0-门诊收费，1-入院登记，2-不区分门诊与住院,3-挂号,4-结帐
    mbytType = bytType
    Me.Show vbModal
    GetPatient = mstrPatient & mstrOther
End Function

Private Sub cmdCancel_Click()
    mstrPatient = ""
    mstrOther = ""
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    '17-门诊起付线支付，18-住院起付线支付，19-本年住院次数，20-门诊费用，21-住院费用，22-帐户余额
    '23-参加统筹支付费用，24-统筹支付费用，25-参加大病支付费用，26-大病支付费用，27-是否特殊参保病人
    '28-参保年限，29-医保状态(0正常)
    If Trim(txtEdit(1).Text) = "" Then
        MsgBox "必须输入病人医保号", vbInformation, gstrSysName
        Exit Sub
    End If
    
    txtEdit(4).Tag = CStr(Year(Date) - CInt(txtEdit(4).Text)) & "-01-01"
    mstrOther = "": mstrPatient = ""
    
    mstrPatient = txtEdit(1).Text & ";"                                 '0 卡号
    mstrPatient = mstrPatient & txtEdit(1).Text & ";"                   '1 医保帐号
    mstrPatient = mstrPatient & ";"                                     '2 密码
    mstrPatient = mstrPatient & txtEdit(2).Text & ";"                   '3 姓名
    mstrPatient = mstrPatient & txtEdit(3).Text & ";"                   '4 性别
    mstrPatient = mstrPatient & txtEdit(4).Tag & ";"                    '5 出生日期
    mstrPatient = mstrPatient & ";"                                     '6 身份证
    mstrPatient = mstrPatient & ";"                                     '7 单位名称/编码
        
    mstrOther = mstrOther & ";"                                         '8 医保机构编码(中心)
    mstrOther = mstrOther & txtEdit(0).Tag & ";"                        '9 顺序号
    mstrOther = mstrOther & ";"                                         '10 身份
    mstrOther = mstrOther & ";"                                         '11 余额
    mstrOther = mstrOther & ";"                                         '12 当前状态
    mstrOther = mstrOther & ";"                                         '13 病种ID
    mstrOther = mstrOther & ";"
    mstrOther = mstrOther & ";"                                         '15 退休证号
    mstrOther = mstrOther & ";"                                         '16 年龄段
    mstrOther = mstrOther & ";"                                         '17 灰度级
    mstrOther = mstrOther & ";"                                         '18 帐户增加累计
    mstrOther = mstrOther & ";"                                         '19 帐户支出累计
    mstrOther = mstrOther & ";"                                         '20 进入统筹累计
    mstrOther = mstrOther & ";"                                         '21 统筹报销累计
    mstrOther = mstrOther & ";"                                         '22 住院次数累计
    mstrOther = mstrOther & ";"                                         '23 就诊类别
    mstrOther = mstrOther & ";"                                         '24 本次起付线
    mstrOther = mstrOther & ";"                                         '25 起付线累计
    mstrOther = mstrOther & ";"                                         '26 基本统筹限额
    
    Me.Hide
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim rsTemp As New ADODB.Recordset
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If Index = 0 Then
        Set rsTemp = gcn徐州农保.Execute("Select * From inPatient Where No='" & Trim(txtEdit(Index)) & "'")
        If rsTemp.EOF Then
            MsgBox "医保前置机中没有该病人的住院信息，请先进行医保入院登记", vbInformation, gstrSysName
            txtEdit(Index).SelStart = 0
            txtEdit(Index).SelLength = Len(txtEdit(Index).Text)
        Else
            txtEdit(0).Tag = rsTemp!ID
            txtEdit(2).Text = Trim(rsTemp!Name)
            txtEdit(3).Text = rsTemp!Sex
            txtEdit(4).Text = rsTemp!age
            If Nvl(rsTemp!id_card, "") = "" Then
                txtEdit(1).Enabled = True
                On Error Resume Next
                DoEvents
                txtEdit(1).SetFocus
                On Error GoTo 0
            Else
                txtEdit(1).Text = rsTemp!id_card
            End If
        End If
    End If
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    txtEdit(Index).SelStart = 0
    txtEdit(Index).SelLength = Len(txtEdit(Index).Text)
End Sub
