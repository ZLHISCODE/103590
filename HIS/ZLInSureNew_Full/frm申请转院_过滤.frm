VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frm申请转院_过滤 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "过滤"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3795
   Icon            =   "frm申请转院_过滤.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   3795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmd取消 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   2400
      TabIndex        =   8
      Top             =   1860
      Width           =   1100
   End
   Begin VB.CommandButton cmd确定 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   1170
      TabIndex        =   7
      Top             =   1860
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Caption         =   "条件(&S)"
      Height          =   1605
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   3435
      Begin VB.ComboBox cbo审核标志 
         Height          =   300
         Left            =   1410
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1080
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker dtp开始日期 
         Height          =   300
         Left            =   1410
         TabIndex        =   2
         Top             =   300
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy年MM月dd日"
         Format          =   64552963
         CurrentDate     =   38063
      End
      Begin MSComCtl2.DTPicker Dtp结束日期 
         Height          =   300
         Left            =   1410
         TabIndex        =   4
         Top             =   690
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy年MM月dd日"
         Format          =   64552963
         CurrentDate     =   38063
      End
      Begin VB.Label lbl审核标志 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "审核标志(&A)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   360
         TabIndex        =   5
         Top             =   1140
         Width           =   990
      End
      Begin VB.Label lbl结束日期 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "结束日期(&E)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   360
         TabIndex        =   3
         Top             =   750
         Width           =   990
      End
      Begin VB.Label lbl开始日期 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "开始日期(&B)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   990
      End
   End
End
Attribute VB_Name = "frm申请转院_过滤"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strStart As String
Public strEnd As String
Public strState As String
Private blnOK As Boolean

Private Sub cmd取消_Click()
    Unload Me
End Sub

Private Sub cmd确定_Click()
    strStart = Format(Me.dtp开始日期.Value, "yyyy-MM-dd")
    strEnd = Format(Me.Dtp结束日期.Value, "yyyy-MM-dd")
    strState = Me.cbo审核标志.ItemData(Me.cbo审核标志.ListIndex)
    If strState = -1 Then strState = "all"
    
    blnOK = True
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlcommfun.PressKey (vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Me.dtp开始日期.Value = Format(DateAdd("d", -10, zldatabase.Currentdate()), "yyyy年MM月DD日")
    Me.Dtp结束日期.Value = Format(zldatabase.Currentdate(), "yyyy年MM月DD日")
    With cbo审核标志
        .Clear
        .AddItem "未审核"
        .ItemData(.NewIndex) = 0
        .AddItem "审核通过"
        .ItemData(.NewIndex) = 1
        .AddItem "审核未通过"
        .ItemData(.NewIndex) = 2
        .AddItem "全部转院申请"
        .ItemData(.NewIndex) = -1
        .ListIndex = 0
    End With
End Sub

Public Function ShowME(str开始日期 As String, str结束日期 As String, str审核标志 As String) As Boolean
    blnOK = False
    
    Me.Show 1
    
    str开始日期 = strStart
    str结束日期 = strEnd
    str审核标志 = strState
    ShowME = blnOK
End Function
