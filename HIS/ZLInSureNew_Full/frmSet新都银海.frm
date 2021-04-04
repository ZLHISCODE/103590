VERSION 5.00
Begin VB.Form frmSet新都银海 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4320
   Icon            =   "frmSet新都银海.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox Txt分中心 
      Enabled         =   0   'False
      Height          =   270
      Left            =   1050
      TabIndex        =   9
      Text            =   "0"
      Top             =   1080
      Width           =   915
   End
   Begin VB.TextBox txtIC端口号 
      Enabled         =   0   'False
      Height          =   315
      Left            =   3818
      TabIndex        =   7
      Text            =   "1"
      Top             =   360
      Width           =   255
   End
   Begin VB.ComboBox cbo卡类型 
      Height          =   300
      Left            =   1043
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   360
      Width           =   2055
   End
   Begin VB.ComboBox cbo适用地区 
      Height          =   300
      ItemData        =   "frmSet新都银海.frx":000C
      Left            =   1043
      List            =   "frmSet新都银海.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1530
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   90
      TabIndex        =   2
      Top             =   2025
      Width           =   4275
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   2255
      TabIndex        =   4
      Top             =   2235
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   965
      TabIndex        =   3
      Top             =   2235
      Width           =   1100
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "IC卡需要指定分中心编号:"
      Height          =   180
      Left            =   180
      TabIndex        =   10
      Top             =   780
      Width           =   2070
   End
   Begin VB.Label lblIC端口号 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "端口号"
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3188
      TabIndex        =   8
      Top             =   420
      Width           =   540
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "卡类型"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   165
      TabIndex        =   6
      Top             =   420
      Width           =   540
   End
   Begin VB.Label lbl适用地区 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "适用地区"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   165
      TabIndex        =   0
      Top             =   1590
      Width           =   720
   End
End
Attribute VB_Name = "frmSet新都银海"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private blnOK As Boolean

Public Function ShowSet() As Boolean
    blnOK = False
    
    Me.Show 1
    ShowSet = blnOK
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cbo卡类型_Click()
    Me.lblIC端口号.Enabled = (cbo卡类型.ListIndex <> 0)
    Me.txtIC端口号.Enabled = (cbo卡类型.ListIndex <> 0)
    Me.Txt分中心.Enabled = (cbo卡类型.ListIndex <> 0)
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHand
    
    gcnOracle.BeginTrans
    gcnOracle.Execute "zl_保险参数_Delete(" & TYPE_新都 & ",NULL)", , adCmdStoredProc
    gcnOracle.Execute "zl_保险参数_Insert(" & TYPE_新都 & ",NULL,'适用地区'," & Me.cbo适用地区.ListIndex & ",1)", , adCmdStoredProc
    gcnOracle.Execute "zl_保险参数_Insert(" & gintInsure & ",NULL,'分中心'," & Me.Txt分中心.Text & ",2)", , adCmdStoredProc
    gcnOracle.CommitTrans
    
    Call SaveSetting("ZLSOFT", "公共模块\" & App.ProductName, "卡类型", Me.cbo卡类型.ListIndex)
    Call SaveSetting("ZLSOFT", "公共模块\" & App.ProductName, "IC设备端口", txtIC端口号.Text)
    
    mint适用地区_新都 = Me.cbo适用地区.ListIndex
    mintIC卡分中心 = Me.Txt分中心.Text
    
    blnOK = True
    Unload Me
    Exit Sub
errHand:
    If ErrCenter = 1 Then Resume
    gcnOracle.RollbackTrans
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    
    '增加初始化数据
    Me.cbo适用地区.Clear
    Me.cbo适用地区.AddItem "新都、青白江"
    '曾明春(2005-12-28):增加适用地区都江堰
    Me.cbo适用地区.AddItem "都江堰"
    Me.cbo适用地区.AddItem "蒲江"
    Me.cbo适用地区.ListIndex = 0
    
    '将以前的参数取出来显示在界面中
    gstrSQL = "Select 参数名,Nvl(参数值,0) Value From 保险参数 Where 险类=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "获取上传入院信息参数值", gintInsure)
    If Not rsTmp.EOF Then
       Select Case rsTmp!参数名
          Case "适用地区"
             Me.cbo适用地区.ListIndex = Nvl(rsTmp!Value, 0)
          Case "分中心"
             Me.Txt分中心.Text = Nvl(rsTmp!Value, 0)
       End Select
    End If

    
    Me.cbo卡类型.Clear
    Me.cbo卡类型.AddItem "磁卡"
    Me.cbo卡类型.AddItem "IC卡-JKP428"
    Me.cbo卡类型.AddItem "IC卡-ICIOX"
    Me.cbo卡类型.ListIndex = 0
    
    cbo卡类型.ListIndex = GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "卡类型", 0)
    txtIC端口号.Text = GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "IC设备端口", 1)

End Sub
