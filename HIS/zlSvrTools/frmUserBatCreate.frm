VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUserBatCreate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "批量创建用户"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5535
   Icon            =   "frmUserBatCreate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4200
      TabIndex        =   3
      Top             =   600
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4200
      TabIndex        =   2
      Top             =   120
      Width           =   1100
   End
   Begin VB.TextBox txtEdit 
      Height          =   350
      Left            =   4200
      MaxLength       =   40
      TabIndex        =   6
      Tag             =   "B.名称"
      Top             =   4080
      Width           =   1100
   End
   Begin VB.CommandButton cmdAllDel 
      Caption         =   "全清(&U)"
      Height          =   350
      Left            =   4200
      TabIndex        =   1
      Top             =   5160
      Width           =   1100
   End
   Begin VB.CommandButton cmdAllSelect 
      Caption         =   "全选(&A)"
      Height          =   350
      Left            =   4200
      TabIndex        =   0
      Top             =   4680
      Width           =   1100
   End
   Begin MSComctlLib.ImageList Img小图标 
      Left            =   4680
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserBatCreate.frx":020A
            Key             =   "Dept"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserBatCreate.frx":0524
            Key             =   "User"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserBatCreate.frx":083E
            Key             =   "Role"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserBatCreate.frx":1118
            Key             =   "Modual"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwDept 
      Height          =   4575
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   8070
      _Version        =   393217
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      Checkboxes      =   -1  'True
      ImageList       =   "Img小图标"
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.Label lblmsg 
      Caption         =   $"frmUserBatCreate.frx":16B2
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label lblEdit 
      AutoSize        =   -1  'True
      Caption         =   "查找(&F)"
      Height          =   180
      Index           =   2
      Left            =   4200
      TabIndex        =   5
      Top             =   3660
      Width           =   630
   End
End
Attribute VB_Name = "frmUserBatCreate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'==============================================================
'==模块变量
'==============================================================
Private mstr所有者 As String
Private mstrDept As String
Private mintOld As Integer
Private mintCount As Integer
'==============================================================
'==公共接口
'==============================================================
Public Function ShowMe(ByVal strOwner As String) As String
    mstr所有者 = strOwner
    mstrDept = ""
    Me.Show 1
    ShowMe = mstrDept
End Function
'==============================================================
'==控件事件
'==============================================================
Private Sub cmdAllDel_Click()
    SelAll False
End Sub

Private Sub cmdAllSelect_Click()
    SelAll True
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer
    
    mstrDept = ""
    For i = 1 To Me.tvwDept.Nodes.Count
        If Me.tvwDept.Nodes(i).Checked Then
            mstrDept = mstrDept & "," & Mid(Me.tvwDept.Nodes(i).Key, 2)
        End If
    Next
    
    If mstrDept = "" Then
        If MsgBox("您尚未选择任何部门，是否继续？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            Unload Me
        End If
    Else
        mstrDept = Mid(mstrDept, 2)
        Unload Me
    End If
End Sub

Private Sub Form_Activate()
    DoEvents
    SelAll True
End Sub

Private Sub Form_Load()
    loadDept
End Sub

Private Sub tvwDept_NodeCheck(ByVal Node As MSComctlLib.Node)
    Dim i As Integer
    
    If Node.Children <> 0 Then
        For i = 1 To Me.tvwDept.Nodes.Count
            If Not Me.tvwDept.Nodes(i).Parent Is Nothing Then
                If Me.tvwDept.Nodes(i).Parent.Checked Then
                    Me.tvwDept.Nodes(i).Checked = True
                ElseIf Not Me.tvwDept.Nodes(i).Parent.Checked And Me.tvwDept.Nodes(i).Text <> Node.Text Then
                    Me.tvwDept.Nodes(i).Checked = False
                End If
            End If
        Next
    End If
End Sub

Private Sub txtEdit_GotFocus()
    Me.txtEdit.SelStart = 0
    Me.txtEdit.SelLength = Len(txtEdit.Text)
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    Dim intOld As Integer
    Dim Node As Node
    Dim blnEnd As Boolean
    
    If KeyAscii <> 13 Then Exit Sub
    If Me.txtEdit.Text = "" Then Exit Sub
    
    Me.txtEdit.SelStart = 0
    Me.txtEdit.SelLength = Len(txtEdit.Text)
    
    If Me.txtEdit.Text <> Me.txtEdit.Tag Then
        intOld = 1
        Me.txtEdit.Tag = Me.txtEdit.Text
        mintOld = 0
    Else
        If mintCount + 1 >= Me.tvwDept.Nodes.Count Then
            Me.txtEdit.Tag = ""
            mintOld = 0
        End If
        intOld = mintOld + 1
    End If
RowX:

    For i = intOld To Me.tvwDept.Nodes.Count
        If InStr(1, Me.tvwDept.Nodes(i).Text, Me.txtEdit.Text) > 0 Or InStr(1, Me.tvwDept.Nodes(i).Tag, UCase(Me.txtEdit.Text)) > 0 Then
            Me.tvwDept.SetFocus
            If Not Me.tvwDept.Nodes(i).Parent Is Nothing Then
                Me.tvwDept.Nodes(i).Parent.Expanded = True
            End If
            Me.tvwDept.Nodes(i).Selected = True
            mintOld = i
            blnEnd = True
            Exit Sub
        End If
    Next
    If Not blnEnd And mintOld <> 0 Then
        mintOld = 0
        intOld = 1
        GoTo RowX
    End If
End Sub
'==============================================================
'==私有方法
'==============================================================
Private Sub SelAll(ByVal blnTemp As Boolean)
    Dim i As Integer
    For i = 1 To Me.tvwDept.Nodes.Count
        Me.tvwDept.Nodes(i).Checked = blnTemp
    Next
End Sub

Private Sub loadDept()
'列出部门表和对应人员
    Dim rsTmp As ADODB.Recordset, strSQL As String
    Dim Node As Node
    
    On Error GoTo errH
    tvwDept.Nodes.Clear
    strSQL = "Select Id, 编码, 名称, 上级id, Zlspellcode(名称) 简码" & vbNewLine & _
        "From " & mstr所有者 & ".部门表" & vbNewLine & _
        "Where 编码 <> '-' And (撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null)" & vbNewLine & _
        "Start With 上级id Is Null" & vbNewLine & _
        "Connect By Prior Id = 上级id"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
    Do While Not rsTmp.EOF
        If IsNull(rsTmp!上级id) Then
            Set Node = tvwDept.Nodes.Add(, , "K" & rsTmp!Id, "【" & rsTmp!编码 & "】" & rsTmp!名称, "Dept", "Dept")
        Else
            Set Node = tvwDept.Nodes.Add("K" & rsTmp!上级id, tvwChild, "K" & rsTmp!Id, "【" & rsTmp!编码 & "】" & rsTmp!名称, "Dept", "Dept")
        End If
        Node.Tag = rsTmp!简码
        rsTmp.MoveNext
    Loop
    Exit Sub
errH:
    MsgBox "[" & err.Number & "]" & err.Description, vbExclamation, Me.Caption
End Sub

