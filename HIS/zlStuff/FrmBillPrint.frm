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

Public Function ShowMe(ByVal FrmParent As Object, ByVal 系统号 As Long, ByVal 票据号 As String, _
                       ByVal 记录状态 As Long, ByVal 单位系数 As Integer, ByVal 单据类型 As Integer, ByVal 单据名称 As String, ByVal 单据号 As String)
    lng系统号 = 系统号
    str票据号 = 票据号
    lng记录状态 = 记录状态
    int单位系数 = 单位系数
    str单据类型 = 单据名称
    int单据类型 = 单据类型
    str单据号 = 单据号
    Me.Show 1, FrmParent
End Function

Private Sub BillPrint(ByVal intPrintMode As Integer)
    Select Case int单据类型
    Case 1712
        ReportOpen gcnOracle, lng系统号, str票据号, Me, "单据编号=" & str单据号, "记录状态=" & lng记录状态, "单位系数=" & int单位系数, str单据类型, intPrintMode
    Case 1713, 1714, 1715, 1716, 1717, 1718, 1719, 1720, 1722, 1725
        ReportOpen gcnOracle, lng系统号, str票据号, Me, "单据编号=" & str单据号, "记录状态=" & lng记录状态, "单位系数=" & int单位系数, intPrintMode
    Case 1721           '
        ReportOpen gcnOracle, lng系统号, str票据号, Me, "单据编号=" & str单据号, "记录状态=" & lng记录状态, intPrintMode
    Case 1724
        ReportOpen gcnOracle, lng系统号, str票据号, Me, "单据编号=" & str单据号, "单位=" & int单位系数, intPrintMode
    End Select
    If intPrintMode <> 1 Then Unload Me
End Sub

Private Sub Form_Load()
    Me.Txt单据类型 = str单据类型
    Me.Txt单据号 = str单据号
End Sub
