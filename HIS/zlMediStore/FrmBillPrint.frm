VERSION 5.00
Begin VB.Form FrmBillPrint 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "方式选择"
   ClientHeight    =   2160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4665
   Icon            =   "FrmBillPrint.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton Cmd输出到Excel 
      Caption         =   "输出到&Excel"
      Height          =   350
      Left            =   2970
      TabIndex        =   7
      Top             =   1710
      Width           =   1425
   End
   Begin VB.CommandButton Cmd预览 
      Caption         =   "预览(&R)"
      Height          =   350
      Left            =   1410
      TabIndex        =   6
      Top             =   1710
      Width           =   1100
   End
   Begin VB.CommandButton Cmd打印 
      Caption         =   "打印(&P)"
      Height          =   350
      Left            =   210
      TabIndex        =   5
      Top             =   1710
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Caption         =   "单据信息"
      Enabled         =   0   'False
      Height          =   1395
      Left            =   210
      TabIndex        =   0
      Top             =   180
      Width           =   4275
      Begin VB.TextBox Txt单据号 
         ForeColor       =   &H80000002&
         Height          =   300
         Left            =   1680
         TabIndex        =   4
         Top             =   810
         Width           =   1935
      End
      Begin VB.TextBox Txt单据类型 
         ForeColor       =   &H80000002&
         Height          =   300
         Left            =   1680
         TabIndex        =   2
         Top             =   390
         Width           =   1935
      End
      Begin VB.Label Lbl单据号 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "单据号(&N)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   750
         TabIndex        =   3
         Top             =   870
         Width           =   810
      End
      Begin VB.Label Lbl单据类型 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "单据类型(&T)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   570
         TabIndex        =   1
         Top             =   450
         Width           =   990
      End
   End
End
Attribute VB_Name = "FrmBillPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lng系统号 As String
Private str票据号 As String
Private str单据类型 As String
Private int单据类型 As Integer
Private str单据号 As String
Private lng记录状态 As Long
Private int单位系数 As Integer
Private mint单据模式 As Integer
Private Sub Cmd打印_Click()
    Call BillPrint(2)
End Sub

Private Sub Cmd输出到Excel_Click()
    Call BillPrint(3)
End Sub

Private Sub Cmd预览_Click()
    Call BillPrint(1)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
End Sub

Public Function ShowME(ByVal frmParent As Object, ByVal 系统号 As Long, ByVal 票据号 As String, _
                       ByVal 记录状态 As Long, ByVal 单位系数 As Integer, ByVal 单据类型 As Integer, ByVal 单据名称 As String, ByVal 单据号 As String, Optional ByVal int单据模式 As Integer = 0)
    lng系统号 = 系统号
    str票据号 = 票据号
    lng记录状态 = 记录状态
    int单位系数 = 单位系数
    str单据类型 = 单据名称
    int单据类型 = 单据类型
    str单据号 = 单据号
    mint单据模式 = int单据模式
    Me.Show 1, frmParent
End Function

Private Sub BillPrint(ByVal intPrintMode As Integer)
    Select Case int单据类型
'    Case 1300           '药品外购入库管理
'        ReportOpen gcnOracle, lng系统号, str票据号, Me, "单据编号=" & str单据号, "记录状态=" & lng记录状态, "单位系数=" & int单位系数, intPrintMode
'    Case 1301           '药品自制入库管理
'        ReportOpen gcnOracle, lng系统号, str票据号, Me, "单据编号=" & str单据号, "记录状态=" & lng记录状态, "单位系数=" & int单位系数, intPrintMode
'    Case 1302           '药品其他入库管理
'        ReportOpen gcnOracle, lng系统号, str票据号, Me, "单据编号=" & str单据号, "记录状态=" & lng记录状态, "单位系数=" & int单位系数, intPrintMode
'    Case 1303           '库存差价调整管理
'        ReportOpen gcnOracle, lng系统号, str票据号, Me, "单据编号=" & str单据号, "记录状态=" & lng记录状态, "单位系数=" & int单位系数, intPrintMode
'    Case 1304           '药品移库管理
'        ReportOpen gcnOracle, lng系统号, str票据号, Me, "单据编号=" & str单据号, "记录状态=" & lng记录状态, "单位系数=" & int单位系数, intPrintMode
'    Case 1305           '药品领用管理
'        ReportOpen gcnOracle, lng系统号, str票据号, Me, "单据编号=" & str单据号, "记录状态=" & lng记录状态, "单位系数=" & int单位系数, intPrintMode
'    Case 1306           '药品其他出库管理
'        ReportOpen gcnOracle, lng系统号, str票据号, Me, "单据编号=" & str单据号, "记录状态=" & lng记录状态, "单位系数=" & int单位系数, intPrintMode
'    Case 1307           '药品盘点管理
'        ReportOpen gcnOracle, lng系统号, str票据号, Me, "单据编号=" & str单据号, "记录状态=" & lng记录状态, "单位系数=" & int单位系数, intPrintMode
    Case 1300
        ReportOpen gcnOracle, lng系统号, str票据号, Me, "单据编号=" & str单据号, "记录状态=" & lng记录状态, "单位系数=" & int单位系数, str单据类型, intPrintMode
    Case 1301, 1302, 1303, 1304, 1305, 1306, 1307, 1344
        ReportOpen gcnOracle, lng系统号, str票据号, Me, "单据编号=" & str单据号, "记录状态=" & lng记录状态, "单位系数=" & int单位系数, intPrintMode
    Case 1320           '药品付款管理
        ReportOpen gcnOracle, lng系统号, str票据号, Me, "单据编号=" & str单据号, "记录状态=" & lng记录状态, intPrintMode
    Case 1330           '药品计划管理
        ReportOpen gcnOracle, lng系统号, str票据号, Me, "单据编号=" & str单据号, IIf(mint单据模式 = 0, "ReportFormat=1", IIf(mint单据模式 = 1, "ReportFormat=2", "ReportFormat=3")), intPrintMode
    End Select
    If intPrintMode <> 1 Then Unload Me
End Sub

Private Sub Form_Load()
    Me.Txt单据类型 = str单据类型
    Me.Txt单据号 = str单据号
End Sub
