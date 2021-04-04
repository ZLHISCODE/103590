VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFeeGroupSelDept 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "缴款组选择"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4560
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   3120
      TabIndex        =   2
      Top             =   2640
      Width           =   1230
   End
   Begin MSComctlLib.ListView lvwMain 
      Height          =   2055
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   3625
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imgMain"
      SmallIcons      =   "imgMain"
      ColHdrIcons     =   "imgMain"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "缴款组名称"
         Object.Width           =   7514
      EndProperty
   End
   Begin MSComctlLib.ImageList imgMain 
      Left            =   2160
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFeeGroupSelDept.frx":0000
            Key             =   "dep"
            Object.Tag             =   "dep"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "请选择财务缴款分组："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   210
      Width           =   2100
   End
End
Attribute VB_Name = "frmFeeGroupSelDept"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub LoadListview()
    '-----------------------------------------------------------------------------------------------------------------------
    '功能:读取当前操作员拥有的缴款组
    '编制:刘尔旋
    '日期:2013-11-07
    '备注:
    '-----------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    Dim lvwItem As ListItem
    On Error GoTo errHandle
    
    strSQL = "Select Id,组名称,负责人ID From 财务缴款分组 Where (删除日期 Is Null or 删除日期 Between Sysdate And to_date('3000-01-01','YYYY-MM-DD')) And 负责人ID=[1]"
    strSQL = strSQL & " Union Select A.组ID,B.组名称,A.组长ID From 财务组组长构成 A,财务缴款分组 B Where A.组ID=B.ID And A.组长ID=[1] And (B.删除日期 Is Null or B.删除日期 Between Sysdate And to_date('3000-01-01','YYYY-MM-DD'))"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID)
    
    Do While Not rsTmp.EOF
        Set lvwItem = lvwMain.ListItems.Add(, "_" & Val(Nvl(rsTmp!ID)), Nvl(rsTmp!组名称), "dep", "dep")
        rsTmp.MoveNext
    Loop
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub cmdOK_Click()
    If lvwMain.SelectedItem.Index = -1 Then Exit Sub
    Dim lngGroupID As Long
    lngGroupID = Val(Mid(lvwMain.SelectedItem.Key, 2))
    '返回缴款组ID至主界面
    Call frmFeeGroupManage.SetGroupID(lngGroupID)
    Unload Me
End Sub

Private Sub Form_Load()
    Call LoadListview
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        cmdOK_Click
    End If
End Sub

Private Sub lvwMain_DblClick()
    cmdOK_Click
End Sub
