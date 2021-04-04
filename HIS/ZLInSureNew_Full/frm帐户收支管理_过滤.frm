VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm帐户收支管理_过滤 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "过滤"
   ClientHeight    =   2940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5745
   Icon            =   "frm帐户收支管理_过滤.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2940
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4410
      TabIndex        =   19
      Top             =   2490
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3180
      TabIndex        =   18
      Top             =   2490
      Width           =   1095
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   210
      TabIndex        =   20
      Top             =   2490
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Caption         =   "条件(&F)"
      Height          =   2295
      Left            =   150
      TabIndex        =   0
      Top             =   90
      Width           =   5445
      Begin VB.CheckBox chk显示过程清单 
         Caption         =   "显示过程清单(&D)"
         Height          =   225
         Left            =   360
         TabIndex        =   17
         Top             =   1920
         Width           =   2415
      End
      Begin VB.TextBox txt经办人 
         Height          =   300
         Left            =   3600
         MaxLength       =   16
         TabIndex        =   16
         Top             =   1500
         Width           =   1605
      End
      Begin VB.ComboBox cbo中心 
         Height          =   300
         Left            =   1050
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1500
         Width           =   1635
      End
      Begin MSComCtl2.DTPicker dtp开始时间 
         Height          =   300
         Left            =   1050
         TabIndex        =   2
         Top             =   330
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy年MM月dd日"
         Format          =   62521347
         CurrentDate     =   37914
      End
      Begin VB.TextBox txt金额_结束 
         Height          =   300
         Left            =   3600
         MaxLength       =   16
         TabIndex        =   12
         Top             =   1110
         Width           =   1605
      End
      Begin VB.TextBox txt金额_开始 
         Height          =   300
         Left            =   1050
         MaxLength       =   16
         TabIndex        =   10
         Top             =   1110
         Width           =   1605
      End
      Begin VB.TextBox txt卡号_结束 
         Height          =   300
         Left            =   3600
         MaxLength       =   20
         TabIndex        =   8
         Top             =   720
         Width           =   1605
      End
      Begin VB.TextBox txt卡号_开始 
         Height          =   300
         Left            =   1050
         MaxLength       =   20
         TabIndex        =   6
         Top             =   720
         Width           =   1605
      End
      Begin MSComCtl2.DTPicker dtp结束时间 
         Height          =   300
         Left            =   3600
         TabIndex        =   4
         Top             =   330
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy年MM月dd日"
         Format          =   62521347
         CurrentDate     =   37914
      End
      Begin VB.Label lbl经办人 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "经办人(&P)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   2760
         TabIndex        =   15
         Top             =   1560
         Width           =   810
      End
      Begin VB.Label lbl医保中心 
         AutoSize        =   -1  'True
         Caption         =   "中心(&R)"
         Height          =   180
         Left            =   330
         TabIndex        =   13
         Top             =   1560
         Width           =   630
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "～"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   3030
         TabIndex        =   3
         Top             =   390
         Width           =   180
      End
      Begin VB.Label lbl时间 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "时间(&T)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   330
         TabIndex        =   1
         Top             =   390
         Width           =   630
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "～"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   2
         Left            =   3030
         TabIndex        =   11
         Top             =   1170
         Width           =   180
      End
      Begin VB.Label lbl金额 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "金额(&M)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   330
         TabIndex        =   9
         Top             =   1170
         Width           =   630
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "～"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   3030
         TabIndex        =   7
         Top             =   780
         Width           =   180
      End
      Begin VB.Label lbl卡号_开始 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "卡号(&A)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   330
         TabIndex        =   5
         Top             =   780
         Width           =   630
      End
   End
End
Attribute VB_Name = "frm帐户收支管理_过滤"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mint险类 As Integer
Private mstrFind As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    mstrFind = ""
    
    '组合查找串
    mstrFind = mstrFind & " And Trunc(B.时间) Between To_Date('" & Format(dtp开始时间.Value, "yyyy-MM-dd") & "','yyyy-MM-dd')" & _
            " And To_Date('" & Format(dtp结束时间.Value, "yyyy-MM-dd") & "','yyyy-MM-dd')"
    If Trim(txt卡号_开始.Text) <> "" Then mstrFind = mstrFind & " And A.卡号>='" & UCase(Trim(txt卡号_开始.Text)) & "'"
    If Trim(txt卡号_结束.Text) <> "" Then mstrFind = mstrFind & " And A.卡号<='" & UCase(Trim(txt卡号_结束.Text)) & "'"
    If Trim(txt金额_开始.Text) <> "" Then mstrFind = mstrFind & " And B.金额>='" & Val(txt金额_开始.Text) & "'"
    If Trim(txt金额_结束.Text) <> "" Then mstrFind = mstrFind & " And B.卡号<='" & Val(txt金额_结束.Text) & "'"
    If cbo中心.ListIndex <> 0 Then mstrFind = mstrFind & " And A.中心=" & cbo中心.ItemData(cbo中心.ListIndex)
    If Trim(txt经办人.Text) <> "" Then mstrFind = mstrFind & " And B.经办人='" & Trim(txt经办人.Text) & "'"
    If chk显示过程清单.Value = 0 Then mstrFind = mstrFind & " And B.性质=1"
    
    Unload Me
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    
    Me.dtp开始时间.Value = Format(DateAdd("m", -1, zlDataBase.Currentdate()), "yyyy年MM月dd日")
    Me.dtp结束时间.Value = Format(zlDataBase.Currentdate(), "yyyy年MM月dd日")
    txt经办人 = gstrUserName
    
    gstrSQL = "Select 名称,序号 ID From 保险中心目录 Where 险类=" & mint险类
    Call OpenRecordset(rsTemp, Me.Caption)
    cbo中心.Clear
    cbo中心.AddItem "所有医保中心"
    cbo中心.ItemData(cbo中心.NewIndex) = 0
    Call zlControl.CboAddData(Me.cbo中心, rsTemp, False)
    Me.cbo中心.ListIndex = 0
End Sub

Public Function ShowME(ByVal frmParent As Object, ByVal int险类 As Integer) As String
    mstrFind = ""
    
    mint险类 = int险类
    Me.Show 1, frmParent
    ShowME = mstrFind
End Function
