VERSION 5.00
Begin VB.Form frmSet成都郊县 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "参数设置"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4320
   Icon            =   "frmSet成都郊县.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox Txt分中心 
      Enabled         =   0   'False
      Height          =   270
      Left            =   960
      TabIndex        =   10
      Text            =   "0"
      Top             =   1380
      Width           =   915
   End
   Begin VB.ComboBox cbo适用地区 
      Height          =   300
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   90
      TabIndex        =   7
      Top             =   2265
      Width           =   4275
   End
   Begin VB.TextBox txtIC端口号 
      Enabled         =   0   'False
      Height          =   315
      Left            =   3810
      TabIndex        =   4
      Text            =   "1"
      Top             =   750
      Width           =   255
   End
   Begin VB.ComboBox cbo卡类型 
      Height          =   300
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   750
      Width           =   2055
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3060
      TabIndex        =   9
      Top             =   2475
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   1770
      TabIndex        =   8
      Top             =   2475
      Width           =   1100
   End
   Begin VB.CheckBox Chk入院信息 
      Caption         =   "入院登记的同时，上传医保病人入院信息(&1)"
      Height          =   345
      Left            =   330
      TabIndex        =   0
      Top             =   240
      Width           =   3855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "IC卡需要指定分中心编号:"
      Height          =   180
      Left            =   180
      TabIndex        =   11
      Top             =   1140
      Width           =   2070
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
      TabIndex        =   5
      Top             =   1860
      Width           =   720
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
      Left            =   3180
      TabIndex        =   3
      Top             =   810
      Width           =   540
   End
   Begin VB.Label lbl卡类型 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "卡类型"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   150
      TabIndex        =   1
      Top             =   810
      Width           =   540
   End
End
Attribute VB_Name = "frmSet成都郊县"
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

Private Sub cbo卡类型_Click()
    Me.lblIC端口号.Enabled = (cbo卡类型.ListIndex <> 0)
    Me.txtIC端口号.Enabled = (cbo卡类型.ListIndex <> 0)
    Me.Txt分中心.Enabled = (cbo卡类型.ListIndex <> 0)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    On Error GoTo errHand
    
    gcnOracle.BeginTrans
    gcnOracle.Execute "zl_保险参数_Delete(" & type_成都郊县 & ",NULL)", , adCmdStoredProc
    gcnOracle.Execute "zl_保险参数_Insert(" & type_成都郊县 & ",NULL,'上传入院信息'," & Chk入院信息.Value & ",1)", , adCmdStoredProc
    gcnOracle.Execute "zl_保险参数_Insert(" & type_成都郊县 & ",NULL,'适用地区'," & Me.cbo适用地区.ListIndex & ",2)", , adCmdStoredProc
    gcnOracle.Execute "zl_保险参数_Insert(" & type_成都郊县 & ",NULL,'分中心'," & Nvl(Me.Txt分中心.Text) & ",3)", , adCmdStoredProc
        
    gcnOracle.CommitTrans
    
    Call SaveSetting("ZLSOFT", "公共模块\" & App.ProductName, "卡类型", Me.cbo卡类型.ListIndex)
    Call SaveSetting("ZLSOFT", "公共模块\" & App.ProductName, "IC设备端口", txtIC端口号.Text)
    
    mint适用地区_成都郊县 = Me.cbo适用地区.ListIndex
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
    Me.cbo适用地区.AddItem "普通郊县"
    Me.cbo适用地区.AddItem "双流县"
    Me.cbo适用地区.AddItem "郫县"
    Me.cbo适用地区.AddItem "温江区"
    Me.cbo适用地区.ListIndex = 0
    
    Me.cbo卡类型.Clear
    Me.cbo卡类型.AddItem "磁卡"
    Me.cbo卡类型.AddItem "IC卡-JKP428"
    Me.cbo卡类型.AddItem "IC卡-ICIOX"
    Me.cbo卡类型.ListIndex = 0
    
    '将以前的参数取出来显示在界面中
    gstrSQL = "Select 参数名,Nvl(参数值,0) Value From 保险参数 Where 险类= [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "获取上传入院信息参数值", 22)
    With rsTmp
        Do While Not rsTmp.EOF
            Select Case !参数名
            Case "上传入院信息"
                Chk入院信息.Value = rsTmp!Value
            Case "适用地区"
                Me.cbo适用地区.ListIndex = rsTmp!Value
            Case "分中心"
                Me.Txt分中心.Text = rsTmp!Value
            End Select
            .MoveNext
        Loop
    End With
    
    cbo卡类型.ListIndex = GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "卡类型", 0)
    txtIC端口号.Text = GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "IC设备端口", 1)
End Sub
