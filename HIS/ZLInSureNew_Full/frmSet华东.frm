VERSION 5.00
Begin VB.Form frmSet华东 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "医保设置"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4965
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CheckBox chk床位费 
      Caption         =   "床位费(&B)"
      Height          =   225
      Left            =   480
      TabIndex        =   3
      Top             =   990
      Width           =   1155
   End
   Begin VB.Frame fra床位费 
      Enabled         =   0   'False
      Height          =   1455
      Left            =   180
      TabIndex        =   4
      Top             =   990
      Width           =   4605
      Begin VB.TextBox txt自费码 
         Height          =   300
         Left            =   3060
         TabIndex        =   10
         Top             =   870
         Width           =   1365
      End
      Begin VB.TextBox txt床位费限额 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   990
         MaxLength       =   3
         TabIndex        =   7
         Top             =   870
         Width           =   705
      End
      Begin VB.Label Label3 
         Caption         =   "    当床位费超过限额后，多余部分对应为自费码，并以新的记录上传"
         ForeColor       =   &H000040C0&
         Height          =   375
         Left            =   300
         TabIndex        =   5
         Top             =   330
         Width           =   4125
      End
      Begin VB.Label lbl床位费自费码 
         AutoSize        =   -1  'True
         Caption         =   "自费码(&Z)"
         Height          =   180
         Left            =   2130
         TabIndex        =   9
         Top             =   930
         Width           =   810
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "元"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   1740
         TabIndex        =   8
         Top             =   930
         Width           =   180
      End
      Begin VB.Label lbl床位费限额 
         AutoSize        =   -1  'True
         Caption         =   "限额(&X)"
         Height          =   180
         Left            =   270
         TabIndex        =   6
         Top             =   930
         Width           =   630
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3690
      TabIndex        =   12
      Top             =   2625
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   2415
      TabIndex        =   11
      Top             =   2625
      Width           =   1100
   End
   Begin VB.CommandButton cmdBrower 
      Caption         =   "浏览(&B)"
      Height          =   350
      Left            =   3840
      TabIndex        =   2
      Top             =   480
      Width           =   945
   End
   Begin VB.TextBox txtPath 
      Height          =   300
      Left            =   180
      TabIndex        =   1
      Top             =   510
      Width           =   3660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "请指定文件存放位置(&L)"
      Height          =   180
      Left            =   180
      TabIndex        =   0
      Top             =   240
      Width           =   1890
   End
End
Attribute VB_Name = "frmSet华东"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng险类 As Long, mblnReturn As Boolean

Public Function ShowME(ByVal lng险类 As Long) As Boolean
    mlng险类 = lng险类
    Me.Show 1
    ShowME = mblnReturn
End Function

Private Sub chk床位费_Click()
    On Error Resume Next
    fra床位费.Enabled = (chk床位费.Value = 1)
    If fra床位费.Enabled Then
        txt床位费限额.SetFocus
    Else
        txt床位费限额.Text = ""
        txt自费码.Text = ""
    End If
End Sub

Private Sub cmdBrower_Click()
    txtPath.Text = BrowPath(Me.hwnd, "请选择文件存放位置：")
End Sub

Private Sub cmdCancel_Click()
    mblnReturn = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Trim(txtPath.Text) = "" Then Exit Sub
    If chk床位费.Value = 1 Then
        If Val(txt床位费限额.Text) <= 0 Then
            MsgBox "床位费限额不能小于等于零！", vbInformation, gstrSysName
            Exit Sub
        End If
        If Trim(txt自费码.Text) = "" Then
            MsgBox "自费码不能为空！", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    gcnOracle.BeginTrans
    On Error GoTo errHand
    
    '删除已经数据
    gstrSQL = "zl_保险参数_Delete(" & mlng险类 & ",null)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '新增参数数据
    gstrSQL = "zl_保险参数_Insert(" & mlng险类 & ",NULL,'文件存放位置','" & txtPath.Text & "',1)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & mlng险类 & ",NULL,'床位费限额','" & txt床位费限额.Text & "',2)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    gstrSQL = "zl_保险参数_Insert(" & mlng险类 & ",NULL,'床位费自费码','" & txt自费码.Text & "',3)"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    mstrSavePath = txtPath.Text
    gcnOracle.CommitTrans
    mblnReturn = True
    
    Unload Me
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    gcnOracle.RollbackTrans
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset, strTemp As String
    gstrSQL = "Select 参数名,参数值 From 保险参数 Where 险类=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, gstrSysName, TYPE_华东)
    Do Until rsTemp.EOF
        strTemp = IIf(IsNull(rsTemp("参数值")), "", rsTemp("参数值"))
        If rsTemp!参数名 = "文件存放位置" Then txtPath.Text = strTemp
        If rsTemp!参数名 = "床位费限额" Then txt床位费限额.Text = strTemp
        If rsTemp!参数名 = "床位费自费码" Then txt自费码.Text = strTemp
        rsTemp.MoveNext
    Loop
    If Val(txt床位费限额.Text) <> 0 Then chk床位费.Value = 1
End Sub
