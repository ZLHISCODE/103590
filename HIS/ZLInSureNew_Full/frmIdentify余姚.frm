VERSION 5.00
Begin VB.Form frmIdentify余姚 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "身份验证"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3900
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   3900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtInfo 
      Height          =   300
      Index           =   4
      Left            =   1095
      TabIndex        =   0
      Top             =   135
      Width           =   2310
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   400
      Left            =   2060
      TabIndex        =   11
      Top             =   2595
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   400
      Left            =   770
      TabIndex        =   10
      Top             =   2595
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   -30
      TabIndex        =   9
      Top             =   2400
      Width           =   3990
   End
   Begin VB.TextBox txtInfo 
      Height          =   300
      Index           =   3
      Left            =   1095
      TabIndex        =   8
      Top             =   1920
      Width           =   2310
   End
   Begin VB.TextBox txtInfo 
      Height          =   300
      Index           =   2
      Left            =   1095
      TabIndex        =   6
      Top             =   1470
      Width           =   2310
   End
   Begin VB.TextBox txtInfo 
      Height          =   300
      Index           =   1
      Left            =   1095
      TabIndex        =   4
      Top             =   1035
      Width           =   2310
   End
   Begin VB.TextBox txtInfo 
      Height          =   300
      Index           =   0
      Left            =   1095
      TabIndex        =   2
      Top             =   585
      Width           =   2310
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "就诊卡号"
      Height          =   180
      Index           =   4
      Left            =   300
      TabIndex        =   12
      Top             =   225
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "单位编号"
      Height          =   180
      Index           =   3
      Left            =   300
      TabIndex        =   7
      Top             =   2010
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "性别"
      Height          =   180
      Index           =   2
      Left            =   660
      TabIndex        =   5
      Top             =   1560
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "姓名"
      Height          =   180
      Index           =   1
      Left            =   660
      TabIndex        =   3
      Top             =   1125
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "个人编号"
      Height          =   180
      Index           =   0
      Left            =   300
      TabIndex        =   1
      Top             =   675
      Width           =   720
   End
End
Attribute VB_Name = "frmIdentify余姚"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbytType As Byte
Private mlng病人ID As Long
Public mstrPatient As String, mstrOther As String

Public Function GetPatient(bytType As Byte, lng病人ID As Long) As String
    '参数：bytType-识别类型，0-门诊收费，1-入院登记，2-不区分门诊与住院,3-挂号,4-结帐
    mbytType = bytType
    mlng病人ID = lng病人ID
    Me.Show vbModal
'    gstrIC明文 = Right(Space(18) & txtInfo(0).Text, 18) & _
'                 String(18, "0") & _
'                 Right(Space(20) & txtInfo(1).Text, 20) & _
'                 Right(Space(2) & txtInfo(2).Text, 2) & _
'                 String(56, "0") & _
'                 Right(Space(10) & txtInfo(3).Text, 10) & _
'                 String(2, "0") & String(126 + 85 + 146 + 116 * 6, "0")
                 
    gstrIC明文 = Space(18 - LenB(StrConv(txtInfo(0).Text, vbFromUnicode))) & txtInfo(0).Text & _
                 String(18, "0") & _
                 Space(20 - LenB(StrConv(txtInfo(1).Text, vbFromUnicode))) & txtInfo(1).Text & _
                 Space(2 - LenB(StrConv(txtInfo(2).Text, vbFromUnicode))) & txtInfo(2).Text & _
                 String(56, "0") & _
                 Space(10 - LenB(StrConv(txtInfo(3).Text, vbFromUnicode))) & txtInfo(3).Text & _
                 String(2, "0") & String(126 + 85 + 146 + 116 * 6, "0")
    GetPatient = mstrPatient & mstrOther
    If GetPatient <> "" Then lng病人ID = mlng病人ID
End Function

Private Sub cmdCancel_Click()
    mstrPatient = "": mstrOther = ""
    Me.Hide
    '取消
End Sub

Private Sub cmdOK_Click()
    '确定
    If txtInfo(2).Text <> "男" And txtInfo(2).Text <> "女" Then
        MsgBox "性别请输入“男”或“女”", vbInformation, gstrSysName
        Exit Sub
    End If
    If Trim(txtInfo(0).Text) = "" Or Trim(txtInfo(1).Text) = "" Or Trim(txtInfo(2).Text) = "" Or Trim(txtInfo(3).Text) = "" Then
        MsgBox "请输入完整的个人身份信息", vbInformation, gstrSysName
        Exit Sub
    End If
    mstrOther = "": mstrPatient = ""
    
    mstrPatient = txtInfo(0).Text & ";"                                 '0 卡号
    mstrPatient = mstrPatient & txtInfo(0).Text & ";"                   '1 医保帐号
    mstrPatient = mstrPatient & ";"                                     '2 密码
    mstrPatient = mstrPatient & txtInfo(1).Text & ";"                   '3 姓名
    mstrPatient = mstrPatient & txtInfo(2).Text & ";"                   '4 性别
    mstrPatient = mstrPatient & ";"                                     '5 出生日期
    mstrPatient = mstrPatient & ";"                                     '6 身份证
    mstrPatient = mstrPatient & "(" & txtInfo(3).Text & ")" & ";"       '7 单位名称/编码
    
    mstrOther = mstrOther & gstr医保机构编码 & ";"                      '8 医保机构编码(中心)
    mstrOther = mstrOther & ";"                                         '9 顺序号
    mstrOther = mstrOther & ";"                                         '10 身份
    mstrOther = mstrOther & ";"                                         '11 余额
    mstrOther = mstrOther & ";"                                         '12 当前状态
    mstrOther = mstrOther & ";"                                         '13 病种ID
    mstrOther = mstrOther & ";"                                         '14 在职状态
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

Private Sub txtInfo_GotFocus(Index As Integer)
    txtInfo(Index).SelStart = 0
    txtInfo(Index).SelLength = Len(txtInfo(Index).Text)
End Sub

Private Sub txtInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim rsTemp As New ADODB.Recordset, rs信息 As New ADODB.Recordset
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If Index = 4 Then
        If Trim(txtInfo(4).Text) <> "" Then
            gstrSQL = "Select * From 病人信息 Where 就诊卡号=[1]"
            Set rs信息 = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CStr(Trim(UCase(txtInfo(4).Text))))
            If Not rs信息.EOF Then
                mlng病人ID = rs信息!病人ID
                gstrSQL = "Select 卡号 From 保险帐户 Where 病人ID=[1]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, CLng(rs信息!病人ID))
                If rsTemp.EOF Then
                    txtInfo(1).Text = rs信息!姓名
                    txtInfo(2).Text = rs信息!性别
                Else
                    txtInfo(0).Text = rsTemp(0)
                    gstrSQL = "Select A.姓名,A.性别,nvl(B.单位编码,'0') From 病人信息 A,保险帐户 B Where A.病人ID=B.病人ID And B.险类=[1] And B.卡号=[2]"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_余姚, CStr(Trim(txtInfo(0).Text)))
                    If Not rsTemp.EOF Then
                        txtInfo(1).Text = rsTemp(0)
                        txtInfo(2).Text = rsTemp(1)
                        txtInfo(3).Text = rsTemp(2)
                    End If
                End If
            End If
        End If
        txtInfo(0).SetFocus
    ElseIf Index = 0 Then
        If Trim(txtInfo(4).Text) = "" Then
            gstrSQL = "Select A.姓名,A.性别,nvl(B.单位编码,'0') From 病人信息 A,保险帐户 B Where A.病人ID=B.病人ID And B.险类=[1] And B.卡号=[2]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_余姚, CStr(Trim(txtInfo(0).Text)))
            If Not rsTemp.EOF Then
                txtInfo(1).Text = rsTemp(0)
                txtInfo(2).Text = rsTemp(1)
                txtInfo(3).Text = rsTemp(2)
            End If
        End If
        txtInfo(1).SetFocus
    ElseIf Index = 3 Then
        cmdOK.SetFocus
    Else
        txtInfo(Index + 1).SetFocus
    End If
End Sub

