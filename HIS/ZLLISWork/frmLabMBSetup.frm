VERSION 5.00
Begin VB.Form frmLabMBSetup 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "参数设置"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   7275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5700
      TabIndex        =   16
      Top             =   2610
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4110
      TabIndex        =   15
      Top             =   2610
      Width           =   1100
   End
   Begin VB.Frame fra计算参数 
      Caption         =   "计算参数"
      Height          =   2325
      Left            =   2820
      TabIndex        =   11
      Top             =   120
      Width           =   4365
      Begin VB.TextBox txt阴性对照 
         Height          =   285
         Left            =   1560
         TabIndex        =   14
         Top             =   698
         Width           =   1185
      End
      Begin VB.CheckBox chk阴性对照 
         Caption         =   "阴性对照小于              时按设定值计算"
         Height          =   180
         Left            =   180
         TabIndex        =   13
         Top             =   750
         Width           =   4065
      End
      Begin VB.CheckBox chk空白对照 
         Caption         =   "是否减去空白对照"
         Height          =   345
         Left            =   180
         TabIndex        =   12
         Top             =   270
         Width           =   2445
      End
   End
   Begin VB.Frame fra通讯参数 
      Caption         =   "通讯参数"
      Height          =   2355
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   2625
      Begin VB.ComboBox cbo通讯口 
         Height          =   300
         ItemData        =   "frmLabMBSetup.frx":0000
         Left            =   1095
         List            =   "frmLabMBSetup.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   1290
      End
      Begin VB.ComboBox cbo波特率 
         Height          =   300
         ItemData        =   "frmLabMBSetup.frx":0004
         Left            =   1095
         List            =   "frmLabMBSetup.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   645
         Width           =   1290
      End
      Begin VB.ComboBox cbo数据位 
         Height          =   300
         ItemData        =   "frmLabMBSetup.frx":0008
         Left            =   1095
         List            =   "frmLabMBSetup.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1050
         Width           =   1290
      End
      Begin VB.ComboBox cbo停止位 
         Height          =   300
         ItemData        =   "frmLabMBSetup.frx":000C
         Left            =   1095
         List            =   "frmLabMBSetup.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1455
         Width           =   1290
      End
      Begin VB.ComboBox cbo校验位 
         Height          =   300
         ItemData        =   "frmLabMBSetup.frx":0010
         Left            =   1095
         List            =   "frmLabMBSetup.frx":0012
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1875
         Width           =   1290
      End
      Begin VB.Label lbl通讯口 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "通讯口(&1)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   240
         TabIndex        =   10
         Top             =   300
         Width           =   810
      End
      Begin VB.Label lbl波特率 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "波特率(&2)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   240
         TabIndex        =   9
         Top             =   705
         Width           =   810
      End
      Begin VB.Label lbl数据位 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "数据位(&3)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   240
         TabIndex        =   8
         Top             =   1110
         Width           =   810
      End
      Begin VB.Label lbl停止位 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "停止位(&4)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   240
         TabIndex        =   7
         Top             =   1515
         Width           =   810
      End
      Begin VB.Label lbl校验位 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "校验位(&5)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   240
         TabIndex        =   6
         Top             =   1935
         Width           =   810
      End
   End
End
Attribute VB_Name = "frmLabMBSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Frame1_DragDrop(Source As Control, x As Single, Y As Single)

End Sub

Private Sub chk阴性对照_Click()
    Me.txt阴性对照.Enabled = Me.chk阴性对照.Value
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    zlDatabase.SetPara "frmLabMB_通讯口", Me.cbo通讯口.Text, 100, 1208
    zlDatabase.SetPara "frmLabMB_波特率", Me.cbo波特率.Text, 100, 1208
    zlDatabase.SetPara "frmLabMB_数据位", Me.cbo数据位.Text, 100, 1208
    zlDatabase.SetPara "frmLabMB_停止位", Me.cbo停止位.Text, 100, 1208
    zlDatabase.SetPara "frmLabMB_校验位", Me.cbo校验位.Text, 100, 1208
    zlDatabase.SetPara "frmLabMB_减空白对照", Me.chk空白对照.Value, 100, 1208
    zlDatabase.SetPara "frmLabMB_阴性对照", Me.chk阴性对照.Value & "," & Me.txt阴性对照.Text, 100, 1208
    Unload Me
End Sub

Private Sub Form_Load()
    Dim aryTemp() As String
    Dim lngCount As Long
    
    '其他固定内容装入
    For lngCount = 1 To 50: Me.cbo通讯口.AddItem "COM" & lngCount: Next
    Me.cbo通讯口.ListIndex = 0

    aryTemp = Split("110|300|600|1200|2400|4800|9600|14400|19200|28800|38400|56000|128000|256000", "|")
    For lngCount = LBound(aryTemp) To UBound(aryTemp): Me.cbo波特率.AddItem aryTemp(lngCount): Next
    Me.cbo波特率.ListIndex = 0

    aryTemp = Split("4|5|6|7|8", "|")
    For lngCount = LBound(aryTemp) To UBound(aryTemp): Me.cbo数据位.AddItem aryTemp(lngCount): Next
    Me.cbo数据位.ListIndex = 0

    aryTemp = Split("1|1.5|2", "|")
    For lngCount = LBound(aryTemp) To UBound(aryTemp):
        Me.cbo停止位.AddItem aryTemp(lngCount):
    Next
    Me.cbo停止位.ListIndex = 0

    aryTemp = Split("E-偶数|M-标记|N-缺省|None|O-奇数|S-空格", "|")
    For lngCount = LBound(aryTemp) To UBound(aryTemp): Me.cbo校验位.AddItem aryTemp(lngCount): Next
    Me.cbo校验位.ListIndex = 0
    
    On Error Resume Next
    
    Me.cbo通讯口 = zlDatabase.GetPara("frmLabMB_通讯口", 100, 1208, "")
    Me.cbo波特率 = zlDatabase.GetPara("frmLabMB_波特率", 100, 1208, "")
    Me.cbo数据位 = zlDatabase.GetPara("frmLabMB_数据位", 100, 1208, "")
    Me.cbo停止位 = zlDatabase.GetPara("frmLabMB_停止位", 100, 1208, "")
    Me.cbo校验位 = zlDatabase.GetPara("frmLabMB_校验位", 100, 1208, "")
    Me.chk空白对照 = zlDatabase.GetPara("frmLabMB_减空白对照", 100, 1208, "0")
    Me.chk阴性对照.Value = Mid(zlDatabase.GetPara("frmLabMB_阴性对照", 100, 1208, "0,"), 1, 1)
    Me.txt阴性对照.Text = Mid(zlDatabase.GetPara("frmLabMB_阴性对照", 100, 1208, "0,"), 3)
End Sub
