VERSION 5.00
Begin VB.Form frm昭通_项目信息 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "项目信息"
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5250
   Icon            =   "frm昭通_项目信息.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3885
      TabIndex        =   14
      Top             =   2205
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   2640
      TabIndex        =   13
      Top             =   2205
      Width           =   1100
   End
   Begin VB.ComboBox cbo医保类型 
      Height          =   300
      Left            =   1020
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   0
      TabIndex        =   12
      Top             =   2040
      Width           =   5715
   End
   Begin VB.Label lbl收费细目 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "收费细目"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   0
      Left            =   180
      TabIndex        =   8
      Top             =   1290
      Width           =   720
   End
   Begin VB.Label lbl收费细目 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "收费细目"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   1
      Left            =   1020
      TabIndex        =   9
      Top             =   1290
      Width           =   3930
   End
   Begin VB.Label lbl辅助信息 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "辅助信息"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   0
      Left            =   180
      TabIndex        =   6
      Top             =   600
      Width           =   720
   End
   Begin VB.Label lbl姓名 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "姓名"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   0
      Left            =   2400
      TabIndex        =   2
      Top             =   300
      Width           =   360
   End
   Begin VB.Label lbl辅助信息 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "医嘱内容"
      ForeColor       =   &H80000008&
      Height          =   540
      Index           =   1
      Left            =   1020
      TabIndex        =   7
      Top             =   600
      Width           =   3930
   End
   Begin VB.Label lbl标识号 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "标识号"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   1
      Left            =   1020
      TabIndex        =   1
      Top             =   300
      Width           =   540
   End
   Begin VB.Label lbl标识号 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "标识号"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   300
      Width           =   540
   End
   Begin VB.Label lbl性别 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "性别"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   1
      Left            =   4650
      TabIndex        =   5
      Top             =   300
      Width           =   360
   End
   Begin VB.Label lbl性别 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "性别"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   0
      Left            =   4140
      TabIndex        =   4
      Top             =   300
      Width           =   360
   End
   Begin VB.Label lbl姓名 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "姓名"
      ForeColor       =   &H80000008&
      Height          =   180
      Index           =   1
      Left            =   2910
      TabIndex        =   3
      Top             =   300
      Width           =   360
   End
   Begin VB.Label lbl医保类型 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "医保类型"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   180
      TabIndex        =   10
      Top             =   1620
      Width           =   720
   End
End
Attribute VB_Name = "frm昭通_项目信息"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mintType As Integer 'intType-调用类型(0-医嘱,1-门诊收费,2-住院记帐)
Private mlng病人ID As Long
Private mlng细目ID As Long
Private mstr摘要 As String
Private mstr备注 As String
Private mbln中草药 As Boolean

Public Function ShowME(ByVal intType As Integer, ByVal lng病人ID As Long, ByVal lng细目ID As Long, _
    ByVal str摘要 As String, ByVal str备注 As String, ByVal bln中草药 As Boolean) As String
    'bln中草药=true,表明是中草药，需选择药品的费用类型；否则是非目录内药品，需选择是大病还是抢救用药
    mintType = intType
    mlng病人ID = lng病人ID
    mlng细目ID = lng细目ID
    mstr摘要 = Trim(UCase(str摘要))
    mstr备注 = str备注
    mbln中草药 = bln中草药
    Me.Show 1
    ShowME = mstr摘要
End Function

Private Sub cmdCancel_Click()
    If mstr摘要 = "" Then mstr摘要 = IIf(mbln中草药, "uzy03", "普通")
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Select Case cbo医保类型.ListIndex
    Case 0
        mstr摘要 = IIf(mbln中草药, "uzy01", "普通")
    Case 1
        mstr摘要 = IIf(mbln中草药, "uzy02", "抢救")
    Case 2
        mstr摘要 = IIf(mbln中草药, "uzy03", "病情需要（大病）")
    End Select
    Unload Me
End Sub

Private Sub Form_Load()
    On Error Resume Next
    Dim rsTemp As New ADODB.Recordset
    
    '读取病人信息
    gstrSQL = "Select 姓名,性别,门诊号,住院号 From 病人信息 Where 病人ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取病人信息", mlng病人ID)
    If mintType = 1 Then
        Me.lbl标识号(1).Caption = Nvl(rsTemp!门诊号)
    Else
        Me.lbl标识号(1).Caption = Nvl(rsTemp!住院号)
    End If
    Me.lbl姓名(1).Caption = Nvl(rsTemp!姓名)
    Me.lbl性别(1).Caption = Nvl(rsTemp!性别)
    
    '显示单据或医嘱的主要信息，如果为空，显示为当前单据
    Me.lbl辅助信息(1).Caption = IIf(Trim(mstr备注) = "", "当前记帐单据", mstr备注)
    
    '提取收费细目信息
    gstrSQL = "Select '品名:('||编码||')'||名称||' 规格:'||Nvl(规格,'') AS 品名 From 收费细目 Where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取收费细目信息", mlng细目ID)
    Me.lbl收费细目(1).Caption = Nvl(rsTemp!品名)
    
    If mbln中草药 Then
        Me.cbo医保类型.AddItem "甲类-uzy01"
        Me.cbo医保类型.AddItem "乙类-uzy02"
        Me.cbo医保类型.AddItem "非医保-uzy03"
    Else
        Me.lbl医保类型.Caption = "用药标识"
        Me.cbo医保类型.AddItem "普通-0"
        Me.cbo医保类型.AddItem "抢救-1"
        Me.cbo医保类型.AddItem "病情需要（大病）-2"
    End If
    Me.cbo医保类型.ListIndex = 0
    
End Sub

Private Sub FindCbo()
    '根据以前设定的摘要，定位当前医保类型
    If mstr摘要 = "" Then Exit Sub
    Select Case mstr摘要
    Case "UZY02", "抢救"
        cbo医保类型.ListIndex = 1
    Case "UZY03", "病情需要（大病）"
        cbo医保类型.ListIndex = 2
    End Select
End Sub
