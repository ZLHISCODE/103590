VERSION 5.00
Begin VB.Form frm贵阳结算信息 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "结算信息"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6735
   Icon            =   "frm贵阳结算信息.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   5220
      TabIndex        =   36
      Top             =   4260
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   -120
      TabIndex        =   37
      Top             =   4050
      Width           =   7125
   End
   Begin VB.TextBox txt单病种编码 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   4920
      TabIndex        =   35
      Top             =   3480
      Width           =   1485
   End
   Begin VB.TextBox txt清算方式 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   4920
      TabIndex        =   33
      Top             =   3090
      Width           =   1485
   End
   Begin VB.TextBox txt个人帐户余额 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   4920
      TabIndex        =   31
      Top             =   2700
      Width           =   1485
   End
   Begin VB.TextBox txt医疗补助支付 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   4920
      TabIndex        =   29
      Top             =   2310
      Width           =   1485
   End
   Begin VB.TextBox txt结算总费用 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   1440
      TabIndex        =   5
      Top             =   1140
      Width           =   1485
   End
   Begin VB.TextBox txt超限额自付 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   4920
      TabIndex        =   27
      Top             =   1920
      Width           =   1485
   End
   Begin VB.TextBox txt个人帐户支付 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   4920
      TabIndex        =   25
      Top             =   1530
      Width           =   1485
   End
   Begin VB.TextBox txt大额统筹自付 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   4920
      TabIndex        =   23
      Top             =   1140
      Width           =   1485
   End
   Begin VB.TextBox txt大额统筹支付 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   4920
      TabIndex        =   21
      Top             =   750
      Width           =   1485
   End
   Begin VB.TextBox txt基本统筹自付 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   4920
      TabIndex        =   19
      Top             =   360
      Width           =   1485
   End
   Begin VB.TextBox txt基本统筹支付 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   1440
      TabIndex        =   17
      Top             =   3480
      Width           =   1485
   End
   Begin VB.TextBox txt进入起付线 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   1440
      TabIndex        =   15
      Top             =   3090
      Width           =   1485
   End
   Begin VB.TextBox txt本次起付线 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   1440
      TabIndex        =   13
      Top             =   2700
      Width           =   1485
   End
   Begin VB.TextBox txt允许报销 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   1440
      TabIndex        =   11
      Top             =   2310
      Width           =   1485
   End
   Begin VB.TextBox txt挂钩自付 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   1440
      TabIndex        =   9
      Top             =   1920
      Width           =   1485
   End
   Begin VB.TextBox txt全自费 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   1440
      TabIndex        =   7
      Top             =   1530
      Width           =   1485
   End
   Begin VB.TextBox txt医院总费用 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   1440
      TabIndex        =   3
      Top             =   750
      Width           =   1485
   End
   Begin VB.TextBox txt医保总费用 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   1440
      TabIndex        =   1
      Top             =   360
      Width           =   1485
   End
   Begin VB.Label lbl单病种编码 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "单病种编码"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3960
      TabIndex        =   34
      Top             =   3540
      Width           =   900
   End
   Begin VB.Label lbl清算方式 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "清算方式"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   4140
      TabIndex        =   32
      Top             =   3150
      Width           =   720
   End
   Begin VB.Label lbl个人帐户余额 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "个人帐户余额"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3780
      TabIndex        =   30
      Top             =   2760
      Width           =   1080
   End
   Begin VB.Label lbl医疗补助支付 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "医疗补助支付"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3780
      TabIndex        =   28
      Top             =   2370
      Width           =   1080
   End
   Begin VB.Label lbl结算总费用 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "结算总费用"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   480
      TabIndex        =   4
      Top             =   1200
      Width           =   900
   End
   Begin VB.Label lbl超限额自付 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "超限额自付"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3960
      TabIndex        =   26
      Top             =   1980
      Width           =   900
   End
   Begin VB.Label lbl个人帐户支付 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "个人帐户支付"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3780
      TabIndex        =   24
      Top             =   1590
      Width           =   1080
   End
   Begin VB.Label lbl大额统筹自付 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "大额统筹自付"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3780
      TabIndex        =   22
      Top             =   1200
      Width           =   1080
   End
   Begin VB.Label lbl大额统筹支付 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "大额统筹支付"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3780
      TabIndex        =   20
      Top             =   810
      Width           =   1080
   End
   Begin VB.Label lbl基本统筹自付 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "基本统筹自付"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3780
      TabIndex        =   18
      Top             =   420
      Width           =   1080
   End
   Begin VB.Label lbl基本统筹支付 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "基本统筹支付"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   300
      TabIndex        =   16
      Top             =   3540
      Width           =   1080
   End
   Begin VB.Label lbl进入起付线 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "进入起付线"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   480
      TabIndex        =   14
      Top             =   3150
      Width           =   900
   End
   Begin VB.Label lbl本次起付线 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "本次起付线"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   480
      TabIndex        =   12
      Top             =   2760
      Width           =   900
   End
   Begin VB.Label lbl允许报销 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "允许报销"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   660
      TabIndex        =   10
      Top             =   2370
      Width           =   720
   End
   Begin VB.Label lbl挂钩自付 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "挂钩自付"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   660
      TabIndex        =   8
      Top             =   1980
      Width           =   720
   End
   Begin VB.Label lbl全自费 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "全自费"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   840
      TabIndex        =   6
      Top             =   1590
      Width           =   540
   End
   Begin VB.Label lbl医院总费用 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "医院总费用"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   480
      TabIndex        =   2
      Top             =   810
      Width           =   900
   End
   Begin VB.Label lbl医保总费用 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "医保总费用"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   480
      TabIndex        =   0
      Top             =   420
      Width           =   900
   End
End
Attribute VB_Name = "frm贵阳结算信息"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    txt本次起付线.Text = Val(Format(GetElemnetValue("STARTFEE"), "#0.00;-#0.00;0;"))
    txt超限额自付.Text = Val(Format(GetElemnetValue("FEEOVER"), "#0.00;-#0.00;0;"))
    txt大额统筹支付.Text = Val(Format(GetElemnetValue("FUND2PAY"), "#0.00;-#0.00;0;"))
    txt大额统筹自付.Text = Val(Format(GetElemnetValue("FUND2SELF"), "#0.00;-#0.00;0;"))
    txt个人帐户余额.Text = Val(Format(GetElemnetValue("ACCTBALANCE"), "#0.00;-#0.00;0;"))
    txt个人帐户支付.Text = Val(Format(GetElemnetValue("ACCTPAY"), "#0.00;-#0.00;0;"))
    txt挂钩自付.Text = Val(Format(GetElemnetValue("FEESELF"), "#0.00;-#0.00;0;"))
    txt基本统筹支付.Text = Val(Format(GetElemnetValue("FUND1PAY"), "#0.00;-#0.00;0;"))
    txt基本统筹自付.Text = Val(Format(GetElemnetValue("FUND1SELF"), "#0.00;-#0.00;0;"))
    txt结算总费用.Text = Val(Format(GetElemnetValue("CALFEEALL"), "#0.00;-#0.00;0;"))
    txt进入起付线.Text = Val(Format(GetElemnetValue("ENTERSTARTFEE"), "#0.00;-#0.00;0;"))
    txt全自费.Text = Val(Format(GetElemnetValue("FEEOUT"), "#0.00;-#0.00;0;"))
    txt医保总费用.Text = Val(Format(GetElemnetValue("FEEALL"), "#0.00;-#0.00;0;"))
    txt医疗补助支付.Text = Val(Format(GetElemnetValue("FUND3PAY"), "#0.00;-#0.00;0;"))
    txt医院总费用.Text = Val(Format(GetElemnetValue("HOSPFEEALL"), "#0.00;-#0.00;0;"))
    txt允许报销.Text = Val(Format(GetElemnetValue("ALLOWFUND"), "#0.00;-#0.00;0;"))
    
    txt单病种编码.Text = GetElemnetValue("SINGLEILLNESSCODE")
    txt清算方式.Text = GetElemnetValue("RECKONINGTYPE")

End Sub
