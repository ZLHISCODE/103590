VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmPeople 
   Caption         =   "操作员信息"
   ClientHeight    =   3180
   ClientLeft      =   5520
   ClientTop       =   5880
   ClientWidth     =   3555
   Icon            =   "frmPeople.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   3555
   StartUpPosition =   1  '所有者中心
   Begin MSComctlLib.ListView lstpeople 
      Height          =   2295
      Left            =   360
      TabIndex        =   14
      Top             =   600
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   4048
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdPeople 
      Caption         =   "…"
      Height          =   285
      Index           =   3
      Left            =   3090
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1845
      Width           =   285
   End
   Begin VB.TextBox txtPeople 
      Enabled         =   0   'False
      Height          =   375
      Index           =   3
      Left            =   960
      TabIndex        =   11
      Top             =   1793
      Width           =   2415
   End
   Begin VB.CommandButton cmdPeople 
      Caption         =   "…"
      Height          =   285
      Index           =   2
      Left            =   3080
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   1325
      Width           =   285
   End
   Begin VB.CommandButton cmdPeople 
      Caption         =   "…"
      Height          =   285
      Index           =   1
      Left            =   3080
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   805
      Width           =   285
   End
   Begin VB.CommandButton cmdPeople 
      Caption         =   "…"
      Height          =   285
      Index           =   0
      Left            =   3080
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   285
      Width           =   285
   End
   Begin VB.CommandButton CmdCancle 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   2280
      TabIndex        =   7
      Top             =   2280
      Width           =   1100
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   840
      TabIndex        =   6
      Top             =   2280
      Width           =   1100
   End
   Begin VB.TextBox txtPeople 
      Enabled         =   0   'False
      Height          =   375
      Index           =   2
      Left            =   960
      TabIndex        =   4
      Top             =   1280
      Width           =   2415
   End
   Begin VB.TextBox txtPeople 
      Enabled         =   0   'False
      Height          =   375
      Index           =   1
      Left            =   960
      TabIndex        =   2
      Top             =   760
      Width           =   2415
   End
   Begin VB.TextBox txtPeople 
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label lblPeople 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "核对人"
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   12
      Top             =   1890
      Width           =   540
   End
   Begin VB.Label lblPeople 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "复核人"
      Height          =   180
      Index           =   3
      Left            =   120
      TabIndex        =   5
      Top             =   1365
      Width           =   540
   End
   Begin VB.Label lblPeople 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "调配人"
      Height          =   180
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   850
      Width           =   540
   End
   Begin VB.Label lblPeople 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "摆药人"
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   330
      Width           =   540
   End
End
Attribute VB_Name = "frmPeople"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mstrReturn As String
Private mlng配制中心id As Long
Private mintIndex As Integer

Private Sub CmdCancle_Click()
    Unload Me
End Sub

Public Function ShowMe(ByVal lng配制中心id As Long) As String
    mstrReturn = ""
    mlng配制中心id = lng配制中心id
    Me.Show 1

    ShowMe = mstrReturn
End Function

Private Sub cmdOk_Click()
    Dim IntCount As Integer
    
    For IntCount = 0 To Me.txtPeople.count - 1
        mstrReturn = mstrReturn & ";" & Me.txtPeople(IntCount).Text
    Next
    
    Unload Me
End Sub

Private Sub loadDate()
    Dim strsql As String
    Dim rsTemp As Recordset
    
    On Error GoTo errHandle
    strsql = "select rownum 行数,A.id,姓名 from 人员表 A,部门人员 B where a.id=b.人员id and b.部门id=[1]  and (撤档时间 is null or 撤档时间=to_date('3000/1/1','yyyy/mm/dd'))"
    Set rsTemp = zlDatabase.OpenSQLRecord(strsql, "", mlng配制中心id)
    
    Do While Not rsTemp.EOF
        lstpeople.ListItems.Add rsTemp!行数, rsTemp!Id & rsTemp!姓名, rsTemp!姓名
        rsTemp.MoveNext
    Loop
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmdPeople_Click(Index As Integer)
    mintIndex = Index
    Me.lstpeople.Visible = True
    Me.lstpeople.SetFocus
End Sub

Private Sub Form_Load()
    Call loadDate
End Sub

Private Sub lstpeople_DblClick()
    Dim IntCount As Integer
    
    Me.txtPeople(mintIndex).Text = ""
    
    For IntCount = 1 To Me.lstpeople.ListItems.count
        If Me.lstpeople.ListItems(IntCount).Checked = True Then
            Me.txtPeople.Item(mintIndex).Text = IIf(Me.txtPeople.Item(mintIndex).Text = "", Me.lstpeople.ListItems(IntCount).Text, Me.txtPeople.Item(mintIndex).Text & "\" & Me.lstpeople.ListItems(IntCount).Text)
        End If
    Next
    
    Me.lstpeople.Visible = False
    For IntCount = 1 To Me.lstpeople.ListItems.count
        Me.lstpeople.ListItems(IntCount).Checked = False
    Next
End Sub

Private Sub lstpeople_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim IntCount As Integer
    
    Me.txtPeople.Item(mintIndex).Text = ""
    If KeyCode = 27 Then
        Me.lstpeople.Visible = False
        
        For IntCount = 1 To Me.lstpeople.ListItems.count
            Me.lstpeople.ListItems(IntCount).Checked = False
        Next
        
    End If
End Sub

