VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmToCollection 
   Caption         =   "添加到收藏"
   ClientHeight    =   5655
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmToCollection.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5655
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox PicButton 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   120
      ScaleHeight     =   615
      ScaleWidth      =   4575
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   4920
      Width           =   4575
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&S)"
         Height          =   360
         Left            =   2400
         TabIndex        =   2
         Top             =   120
         Width           =   990
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "取消(&C)"
         Height          =   360
         Left            =   3600
         TabIndex        =   1
         Top             =   120
         Width           =   990
      End
   End
   Begin MSComctlLib.TreeView tvwCollectionType 
      Height          =   4095
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   7223
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "imgList"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   4080
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmToCollection.frx":6852
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmToCollection.frx":6BEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmToCollection.frx":6F86
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image imgTree 
      Height          =   480
      Index           =   1
      Left            =   240
      Picture         =   "frmToCollection.frx":7320
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblPrompt 
      AutoSize        =   -1  'True
      Caption         =   "请选择需要收藏到的目录"
      Height          =   195
      Left            =   960
      TabIndex        =   4
      Top             =   240
      Width           =   1980
   End
End
Attribute VB_Name = "frmToCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private strSQL As String
Private mAdviceID As Long
Private mSendID As Long

Public Sub ShowToCollectionWind(Optional owner As Form = Nothing, Optional AdviceId As Long, Optional SendID As Long)
'显示收藏管理窗口
    mAdviceID = AdviceId
    mSendID = SendID
    
    '加载TreeView数据
    Call LoadTreeView
    If tvwCollectionType.Nodes.Item(1).Children = 0 Then
        MsgBox "请先到收藏管理中增加收藏目录", , gstrSysName
        Exit Sub
    End If
 
    Call Me.Show(1, owner)
End Sub

Private Sub LoadTreeView()
'加载TreeView数据方法
    Dim i As Long
    Dim objNode As Node
    Dim strSQL As String
    Dim rsTvwData As ADODB.Recordset
    
On Error GoTo errHand

    strSql = "select ID,上级ID,收藏类别,是否共享 from 影像收藏类别 where 创建人ID= " & UserInfo.ID & " or 创建人ID is null Start With 上级id Is Null Connect By Prior ID = 上级id"
    Set rsTvwData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    With rsTvwData
        Me.tvwCollectionType.Nodes.Clear
        
        Do While Not .EOF
            If IsNull(!上级ID) Then
                Set objNode = Me.tvwCollectionType.Nodes.Add(, , "_" & Nvl(!ID), Nvl(!收藏类别), IIf(!是否共享 = 0, 1, 3), IIf(Nvl(!是否共享) = 0, 2, 3))
            Else
                Set objNode = Me.tvwCollectionType.Nodes.Add("_" & Nvl(!上级ID), tvwChild, "_" & Nvl(!ID), Nvl(!收藏类别), IIf(Nvl(!是否共享) = 0, 1, 3), IIf(Nvl(!是否共享) = 0, 2, 3))
            End If
            objNode.Sorted = True
            objNode.Expanded = True
            .MoveNext
        Loop
    End With
    
   Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub tvwCollectionType_NodeClick(ByVal Node As MSComctlLib.Node)
On Error GoTo errHand
'如果为顶级节点，则禁用确定按钮
    
 If Trim(Node.Text) = "收藏类别" Then
    cmdOK.Enabled = False
 Else
    cmdOK.Enabled = True
 End If
    
 Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub CmdOK_Click()
'执行添加收藏操作
On Error GoTo errHand
Dim rsTemp As ADODB.Recordset
Dim dtServicesTime As String
Dim strSQL As String


    
     '判断相同收藏类型下 收藏内容是否重复
     strSql = "select b.医嘱id from 影像收藏类别 a,影像收藏内容 b where a.id = b.收藏id and a.创建人ID= " & UserInfo.ID & " and a.收藏类别='" & Trim(tvwCollectionType.SelectedItem.Text) & "'"
     
     Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
     
     Do While Not rsTemp.EOF
        If Nvl(rsTemp!医嘱ID) = mAdviceID Then
            Call MsgBoxD(Me, "该检查已被[ " & Trim(tvwCollectionType.SelectedItem.Text) & " ]收藏。", vbOKOnly, Me.Caption)
            Exit Sub
        End If
        rsTemp.MoveNext
     Loop
    
    '当前服务器时间
    dtServicesTime = zlDatabase.Currentdate
     
    strSQL = "Zl_影像收藏内容_新增(" & Val(Mid(tvwCollectionType.SelectedItem.Key, 2, 5)) & "," & mAdviceID & "," & zlStr.To_Date(dtServicesTime) & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
     
     '添加成功 关闭窗口
     Unload Me

 Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdCancel_Click()
'关闭窗口
On Error GoTo errHand

    Unload Me

 Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    
    '初始化时禁用确定按钮
    cmdOK.Enabled = False
    
End Sub

Private Sub Form_Resize()
On Error Resume Next
    tvwCollectionType.Top = 660
    tvwCollectionType.Left = 80
    tvwCollectionType.Height = Me.ScaleHeight - PicButton.Height - 120
    tvwCollectionType.Width = Me.ScaleWidth - 160

    PicButton.Top = tvwCollectionType.Height + 140
    PicButton.Left = 0
    PicButton.Width = Me.ScaleWidth

    cmdOK.Left = PicButton.Width - cmdOK.Width - 1300
    cmdCancel.Left = PicButton.Width - cmdCancel.Width - 100
    
End Sub




