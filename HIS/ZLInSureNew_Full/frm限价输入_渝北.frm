VERSION 5.00
Begin VB.Form frm限价输入_渝北 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "限价审批信息输入"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5940
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4635
      TabIndex        =   6
      Top             =   3315
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   3210
      TabIndex        =   5
      Top             =   3300
      Width           =   1100
   End
   Begin VB.Frame fra 
      Height          =   45
      Index           =   0
      Left            =   -1095
      TabIndex        =   8
      Top             =   3150
      Width           =   9300
   End
   Begin VB.Frame fra 
      Height          =   45
      Index           =   1
      Left            =   -45
      TabIndex        =   7
      Top             =   630
      Width           =   6135
   End
   Begin VB.CheckBox chk审批 
      Caption         =   "审批通过"
      Height          =   240
      Left            =   1260
      TabIndex        =   0
      Top             =   2370
      Value           =   1  'Checked
      Width           =   1080
   End
   Begin VB.TextBox txtEdit 
      Height          =   315
      Index           =   0
      Left            =   1260
      TabIndex        =   2
      Top             =   2693
      Width           =   1530
   End
   Begin VB.TextBox txtEdit 
      Height          =   315
      Index           =   1
      Left            =   4245
      TabIndex        =   4
      Top             =   2693
      Width           =   1575
   End
   Begin VB.Image img 
      Height          =   555
      Left            =   120
      Picture         =   "frm限价输入_渝北.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   615
   End
   Begin VB.Label lbl 
      Caption         =   "医保项目价格大于200.00元需确定相关的审批信息。"
      Height          =   225
      Index           =   0
      Left            =   945
      TabIndex        =   18
      Top             =   135
      Width           =   4965
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "收费项目编码"
      Height          =   180
      Index           =   2
      Left            =   120
      TabIndex        =   17
      Top             =   855
      Width           =   1080
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   0
      Left            =   1260
      TabIndex        =   16
      Top             =   795
      Width           =   1065
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "收费项目名称"
      Height          =   180
      Index           =   3
      Left            =   120
      TabIndex        =   15
      Top             =   1245
      Width           =   1080
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   1
      Left            =   1260
      TabIndex        =   14
      Top             =   1185
      Width           =   4560
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "医保项目编码"
      Height          =   180
      Index           =   4
      Left            =   120
      TabIndex        =   13
      Top             =   1635
      Width           =   1080
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   2
      Left            =   1260
      TabIndex        =   12
      Top             =   1575
      Width           =   1065
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "医保项目名称"
      Height          =   180
      Index           =   5
      Left            =   120
      TabIndex        =   11
      Top             =   2040
      Width           =   1080
   End
   Begin VB.Label lblEdit 
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Index           =   3
      Left            =   1260
      TabIndex        =   10
      Top             =   1980
      Width           =   4560
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "审批医生"
      Height          =   180
      Index           =   1
      Left            =   480
      TabIndex        =   1
      Top             =   2760
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "审批医生职称"
      Height          =   180
      Index           =   6
      Left            =   3105
      TabIndex        =   3
      Top             =   2760
      Width           =   1080
   End
   Begin VB.Label lblInfor 
      Height          =   150
      Left            =   975
      TabIndex        =   9
      Top             =   390
      Width           =   5040
   End
End
Attribute VB_Name = "frm限价输入_渝北"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mstrVerify As String
Dim mlng费用ID As Long
 
Dim mstrCode As String
Private Sub chk审批_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
 
End Sub

Private Sub cmdCancel_Click()
    mstrVerify = ""
    Unload Me
End Sub

Private Sub cmdOK_Click()
   Dim strSQL As String
    mstrVerify = chk审批.Value & "||" & txtEdit(0).Text & "||" & txtEdit(1).Text
    '--ZL_医保明细费用_INSERT(
    '    费用ID_IN IN 医保明细费用.费用ID%TYPE,
    '    审批人_IN IN 医保明细费用.审批人%TYPE,
    '    审批人职称_IN IN 医保明细费用.审批人职称%TYPE,
    '    审批标志_IN IN 医保明细费用.审批标志%TYPE
    '   就诊编号,结算编号
    '    退单记帐流水号
    strSQL = "ZL_医保明细费用_INSERT(" & _
         mlng费用ID & "," & _
         "'" & txtEdit(0).Text & "'," & _
         "'" & txtEdit(1).Text & "'," & _
         IIf(chk审批.Value = 1, 2, 0) & ",'" & _
         g病人身份_重庆渝北.就诊编号 & "','" & _
         g病人身份_重庆渝北.结算编号 & "'," & _
         "NULL," & _
         IIf(mstrCode = "", "NULL", "'" & mstrCode & "'") & ")"
    
    Call SQLTest(App.ProductName, Me.Caption, strSQL)
    gcnOracle_CQYB.Execute strSQL, , adCmdStoredProc
    Call SQLTest
    Unload Me
End Sub

Private Sub Form_Load()
    lbl(0).Caption = "医保项目价格大于" & InitInfor_重庆渝北.单价限价 & "元需确定相关的审批信息。"
    Me.cmdOK.Enabled = True
End Sub
Public Function Get审批信息(lng费用ID As Long, strCode As String) As String
    Dim rsTemp As New ADODB.Recordset
        
    gstrSQL = " select A.ID,NO,序号, b.编码,b.名称,c.项目编码, c.项目名称 " & _
             " from 门诊费用记录 a, 收费细目 b,保险支付项目 c " & _
             " where a.收费细目id=b.id and a.收费细目id=c.收费细目id and a.id=[1]" & _
             " UNION " & _
             " select A.ID,NO,序号, b.编码,b.名称,c.项目编码, c.项目名称 " & _
             " from 住院费用记录 a, 收费细目 b,保险支付项目 c " & _
             " where a.收费细目id=b.id and a.收费细目id=c.收费细目id and a.id=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "获取相关明细记录信息!", lng费用ID)
    
    If rsTemp.EOF Then
        Exit Function
    End If
    
    lblInfor.Caption = "NO:" & Nvl(rsTemp!NO) & "   行号:" & Nvl(rsTemp!序号, 0)
    
    lblEdit(0).Caption = Nvl(rsTemp!编码)
    lblEdit(1).Caption = Nvl(rsTemp!编码)
    lblEdit(2).Caption = Nvl(rsTemp!项目编码)
    lblEdit(3).Caption = Nvl(rsTemp!项目名称)
    mlng费用ID = lng费用ID
    mstrCode = strCode
    Me.Show 1
    Get审批信息 = mstrVerify
End Function

Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    End If
End Sub
