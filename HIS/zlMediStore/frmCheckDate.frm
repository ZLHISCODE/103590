VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCheckDate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "盘点条件设置"
   ClientHeight    =   1875
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   Icon            =   "frmCheckDate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1875
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   300
      Left            =   3120
      TabIndex        =   16
      Top             =   1440
      Width           =   1100
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "确定(&O)"
      Height          =   300
      Left            =   1950
      TabIndex        =   0
      Top             =   1440
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   300
      Left            =   360
      TabIndex        =   17
      Top             =   1440
      Width           =   1100
   End
   Begin VB.Frame fraCondition 
      Caption         =   "条件"
      Height          =   1200
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4335
      Begin VB.PictureBox pic毒麻精神 
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   960
         ScaleHeight     =   615
         ScaleWidth      =   3015
         TabIndex        =   11
         Top             =   1800
         Width           =   3015
         Begin VB.CheckBox chk药品类型 
            Caption         =   "精神II类"
            Height          =   180
            Index           =   3
            Left            =   1440
            TabIndex        =   15
            Top             =   360
            Width           =   1215
         End
         Begin VB.CheckBox chk药品类型 
            Caption         =   "毒性药"
            Height          =   180
            Index           =   2
            Left            =   120
            TabIndex        =   14
            Top             =   360
            Width           =   855
         End
         Begin VB.CheckBox chk药品类型 
            Caption         =   "精神I类"
            Height          =   180
            Index           =   1
            Left            =   1440
            TabIndex        =   13
            Top             =   120
            Width           =   1215
         End
         Begin VB.CheckBox chk药品类型 
            Caption         =   "麻醉药"
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   12
            Top             =   120
            Value           =   1  'Checked
            Width           =   855
         End
      End
      Begin VB.PictureBox pic近效期 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   960
         ScaleHeight     =   495
         ScaleWidth      =   3015
         TabIndex        =   6
         Top             =   1080
         Width           =   3015
         Begin VB.TextBox txt 
            BackColor       =   &H8000000E&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   1440
            TabIndex        =   9
            Text            =   "30"
            Top             =   270
            Width           =   300
         End
         Begin VB.CheckBox chkDay 
            Caption         =   "    天"
            Height          =   255
            Left            =   1200
            TabIndex        =   8
            Top             =   240
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox chk失效 
            Caption         =   "已失效"
            Height          =   255
            Left            =   2160
            TabIndex        =   10
            Top             =   240
            Width           =   855
         End
         Begin VB.Label lbl效期 
            AutoSize        =   -1  'True
            Caption         =   "近效期时间"
            Height          =   180
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   900
         End
      End
      Begin VB.ComboBox cboType 
         Height          =   300
         Left            =   1080
         TabIndex        =   5
         Text            =   "Combo1"
         Top             =   720
         Width           =   3015
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Left            =   1080
         TabIndex        =   3
         Top             =   240
         Width           =   3045
         _ExtentX        =   5371
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy年MM月dd日 HH:mm:ss"
         Format          =   185794563
         CurrentDate     =   36901
      End
      Begin VB.Label lblType 
         AutoSize        =   -1  'True
         Caption         =   "盘点类型"
         Height          =   180
         Left            =   240
         TabIndex        =   4
         Top             =   780
         Width           =   720
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "盘点时间"
         Height          =   180
         Left            =   240
         TabIndex        =   2
         Top             =   300
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmCheckDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnReturn As Boolean
Private mint盘点时间范围 As Integer
Private mstr盘点时间 As String
Private mint编辑状态 As Integer             '7、库房全部药品盘点；8、特殊药品盘点；9、自动生成有账面数量未盘点的药品
Private mint盘点类型 As Integer
Private mstr近效期 As String
Private mstr毒理 As String

Public Function GetCondition(FrmMain As Form, ByRef str盘点时间 As String, ByVal int编辑状态 As Integer, ByRef int盘点类型 As Integer, ByRef str近效期 As String, ByRef str毒理 As String) As Boolean
    
    mblnReturn = False
    mint编辑状态 = int编辑状态
    
    Me.Show 1, FrmMain
    
    str盘点时间 = mstr盘点时间
    int盘点类型 = mint盘点类型
    str近效期 = mstr近效期
    str毒理 = mstr毒理
    
    GetCondition = mblnReturn
    
End Function


Private Sub cboType_Click()
    '选择近效期可见
    pic近效期.Visible = cboType.ListIndex = 0
    If pic近效期.Visible Then pic近效期.Top = 1080
    '选择毒麻精神可见
    pic毒麻精神.Visible = cboType.ListIndex = 1
    If pic毒麻精神.Visible Then pic毒麻精神.Top = 1080
    
    If cboType.ListIndex > 1 Then '调整窗体
        fraCondition.Height = 1200
        cmdHelp.Top = fraCondition.Top + fraCondition.Height + 120
        CmdSave.Top = cmdHelp.Top
        CmdCancel.Top = cmdHelp.Top
        Me.Height = cmdHelp.Top + cmdHelp.Height + 520 '改变窗体高度
    Else
        fraCondition.Height = 1800
        cmdHelp.Top = fraCondition.Top + fraCondition.Height + 120
        CmdSave.Top = cmdHelp.Top
        CmdCancel.Top = cmdHelp.Top
        Me.Height = cmdHelp.Top + cmdHelp.Height + 520 '改变窗体高度
    End If
    
End Sub


Private Sub CmdCancel_Click()
    mblnReturn = False
    Unload Me
End Sub

Private Sub CmdSave_Click()
    mblnReturn = True
    mstr盘点时间 = Format(dtpDate.Value, "yyyy-MM-dd hh:mm:ss")
    If cboType.ListIndex = 0 And chkDay.Value = 0 And chk失效.Value = 0 Then
        MsgBox "请对近效期时间进行设置！", vbInformation + vbOKOnly, gstrSysName
        chkDay.SetFocus
        Exit Sub
    ElseIf cboType.ListIndex = 1 And (chk药品类型(0).Value = 0 And chk药品类型(1).Value = 0 And chk药品类型(2).Value = 0 And chk药品类型(3).Value = 0) Then
        MsgBox "请选择盘点药品毒理类型！", vbInformation + vbOKOnly, gstrSysName
        chk药品类型(0).SetFocus
        Exit Sub
    End If
    
    mint盘点类型 = cboType.ListIndex
    mstr近效期 = IIf(chkDay.Value = 0, 0, Val(txt.Text)) & ":" & chk失效.Value '近效期返回条件
    mstr毒理 = chk药品类型(0).Value & ":" & chk药品类型(2).Value & ":" & chk药品类型(1).Value & ":" & chk药品类型(3).Value '返回顺序是麻醉、毒性、精神I类、精神II类
    
    Unload Me
End Sub

Private Sub dtpDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_Load()
    mint盘点时间范围 = Val(zlDataBase.GetPara("盘点时间范围设置", glngSys, 1307, 30))
    dtpDate.MinDate = CDate(Format(DateAdd("d", -mint盘点时间范围, Date), "yyyy-mm-dd") & " 00:00:00")
    '药品材质权限控制
    
    dtpDate.Value = Format(Sys.Currentdate, dtpDate.CustomFormat)
    dtpDate.MaxDate = dtpDate.Value
    
    If mint编辑状态 = 8 Then '特殊药品盘点
        cboType.AddItem "0-近效期药品"
        cboType.AddItem "1-毒麻精神药品"
        cboType.AddItem "2-停用药品"
'        cboType.AddItem "3-无库存记录的药品"
        cboType.AddItem "3-无数量但有库存金额或差价的药品"
        cboType.AddItem "4-基本药物"
        
        cboType.ListIndex = 0
    ElseIf mint编辑状态 = 7 Or mint编辑状态 = 9 Then
        fraCondition.Height = 720
        cmdHelp.Top = fraCondition.Top + fraCondition.Height + 120
        CmdSave.Top = cmdHelp.Top
        CmdCancel.Top = cmdHelp.Top
        Me.Height = cmdHelp.Top + cmdHelp.Height + 520 '改变窗体高度
    End If
End Sub


Private Sub txt_Change()
    If (Val(txt.Text) <= 0 Or Val(txt.Text) > 999) And txt.Text <> "" Then
        txt.Text = 30
    End If
End Sub

Private Sub txt_KeyPress(KeyAscii As Integer)
    If InStr("0123456789", Chr(KeyAscii)) > 0 Or KeyAscii = 8 Then
        KeyAscii = KeyAscii
    Else
        KeyAscii = 0
    End If
End Sub
