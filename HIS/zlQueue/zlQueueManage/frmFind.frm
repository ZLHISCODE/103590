VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFind 
   Caption         =   "查找"
   ClientHeight    =   5268
   ClientLeft      =   72
   ClientTop       =   360
   ClientWidth     =   8580
   Icon            =   "frmFind.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5268
   ScaleWidth      =   8580
   Begin VB.PictureBox Picture2 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      ScaleHeight     =   852
      ScaleWidth      =   8580
      TabIndex        =   7
      Top             =   4410
      Width           =   8580
      Begin VB.CommandButton cmdExit 
         Caption         =   "退出(&Q)"
         Height          =   375
         Index           =   0
         Left            =   7080
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "回诊(&C)"
         Height          =   375
         Index           =   1
         Left            =   4680
         TabIndex        =   9
         Top             =   240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "恢复(&R)"
         Height          =   375
         Index           =   2
         Left            =   5880
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   0
      ScaleHeight     =   972
      ScaleWidth      =   8580
      TabIndex        =   1
      Top             =   0
      Width           =   8580
      Begin VB.ComboBox cboFindWay 
         Height          =   300
         ItemData        =   "frmFind.frx":000C
         Left            =   1080
         List            =   "frmFind.frx":001C
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   120
         Width           =   2655
      End
      Begin VB.TextBox txtFindData 
         Height          =   300
         Left            =   1080
         TabIndex        =   3
         Top             =   555
         Width           =   2655
      End
      Begin VB.CommandButton cmdStartFind 
         Caption         =   "开始查找(&F)"
         Height          =   375
         Left            =   3840
         TabIndex        =   2
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label labFindWay 
         Caption         =   "查找方式："
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   165
         Width           =   975
      End
      Begin VB.Label labFindData 
         Caption         =   "查找数据："
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   975
      End
   End
   Begin MSComctlLib.ListView lvwQueueData 
      Height          =   3255
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   8055
      _ExtentX        =   14203
      _ExtentY        =   5736
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "患者姓名"
         Text            =   "患者姓名"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "排队号码"
         Text            =   "排队号码"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "科室名称"
         Text            =   "科室名称"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "诊室名称"
         Text            =   "诊室名称"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "医生姓名"
         Text            =   "医生姓名"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Key             =   "队列名称"
         Text            =   "队列名称"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Key             =   "回诊序号"
         Text            =   "回诊序号"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Key             =   "排队时间"
         Text            =   "排队时间"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Key             =   "排队状态"
         Text            =   "排队状态"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mcnOracle As ADODB.Connection  'oracle 数据库连接
Private mstrFindKey As String          '数据的查找方式
Private gbyt磁卡 As Byte               '磁卡号长度

Public Sub ShowFind(cnOracle As ADODB.Connection, ByVal lngCardLen As Long, Optional owner As Form = Null)
    Set mcnOracle = cnOracle
    
    gbyt磁卡 = lngCardLen
    Me.Show 1, owner
End Sub



Private Sub cmdExit_Click(Index As Integer)
    Dim strQueueId As String
    
    On Error GoTo errHandle
    
    Select Case Index
        Case 0
            Unload Me
        Case 1, 2
            strQueueId = GetSelectId()
          
            If Trim(strQueueId) = "" Then
                MsgBox "尚未选择一条需要进行复诊操作的数据。", vbInformation, "排队叫号系统"
                Exit Sub
            End If
            
            If Index = 1 Then
                Call Execute_回诊(Val(strQueueId))
            ElseIf Index = 2 Then
                Call Execute_恢复(Val(strQueueId))
            End If
            
            '刷新数据
            Call cmdStartFind_Click
            
            'MsgBox "操作执行完成。", vbInformation, "排队叫号系统"
    End Select
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub


Private Function GetSelectId() As String
'***************************************
'
'取得当前选中的数据
'
'***************************************
    On Error GoTo errHandle
        
        If lvwQueueData.SelectedItem Is Nothing Then
          GetSelectId = ""
          Exit Function
        End If
        
        GetSelectId = lvwQueueData.SelectedItem.Tag
        
    Exit Function
errHandle:
      GetSelectId = ""
      If ErrCenter = 1 Then Resume
End Function


Private Sub cmdStartFind_Click()
    Dim rsData As ADODB.Recordset
    Dim strFindType As String
    Dim strFindValue As String
    
    On Error GoTo errHandle
    strFindValue = txtFindData.Text
    
    If Trim(strFindValue) = "" Then
        MsgBox "请输入需要查找的数据值。", vbOKOnly, Me.Caption
        
        Call txtFindData.SetFocus
        Exit Sub
    End If
    
    Call lvwQueueData.ListItems.Clear
    
    '取得检索类型
    strFindType = cboFindWay.Text
    
    Set rsData = FindQueueData(strFindType, strFindValue)
    
    If rsData Is Nothing Then
        MsgBox "没有检索到所需数据。", vbInformation, "排队叫号系统"
        Exit Sub
    End If
    
    If rsData.RecordCount <= 0 Then
        MsgBox "没有检索到所需数据。", vbInformation, "排队叫号系统"
        Exit Sub
    End If
    
    Call LoadDataToFace(lvwQueueData, rsData, "ID")
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub



Private Function FindQueueData(ByVal findType As String, ByVal findData As String) As ADODB.Recordset
    Dim strSql As String, strFilter As String
    Dim str门诊号 As String, str姓名 As String, str就诊卡号 As String, str医保号 As String, str挂号单号 As String
    
    On Error GoTo errHandle
    
    strFilter = ""
    
    Select Case findType  ' '0-门诊号;1-姓名;2-挂号单;3-就诊卡号;4-医保号
    Case "门诊号"
        str门诊号 = Val(findData)
        strFilter = strFilter & " And A.门诊号 = [1]"
    Case "姓名"
        str姓名 = findData & "%"
        strFilter = strFilter & " And A.姓名 Like [2]"
    Case "就诊卡号"
        str就诊卡号 = findData
        strFilter = strFilter & " And A.就诊卡号=[3]"
    Case Else    ' "医保号"
        str医保号 = findData
        strFilter = strFilter & " And A.医保号=[4]"
    End Select
    
            
    If Trim(findType) <> "姓名" Then
        strSql = "Select q.ID, q.队列名称, p.名称 as 科室名称, q.患者姓名, q.排队号码, q.诊室 as 诊室名称, " & _
                 " q.医生姓名, q.回诊序号, q.排队时间, decode(q.排队状态, 1, '呼叫中', 0, '排队中', 3, '暂停', 4, '完成', '已弃号') as 排队状态  " & vbCrLf & _
                 " From 病人信息 A, 排队叫号队列 Q, 部门表 P " & vbCrLf & _
                 " Where Q.病人id = A.病人ID and Q.科室ID=P.ID " & vbCrLf & strFilter
    Else
        strSql = "Select q.ID, q.队列名称, p.名称 as 科室名称, q.患者姓名, q.排队号码, q.诊室 as 诊室名称, " & _
                 " q.医生姓名, q.回诊序号, q.排队时间, decode(q.排队状态, 1, '呼叫中', 0, '排队中', 3, '暂停', 4, '完成', '已弃号') as 排队状态  " & vbCrLf & _
                 " From 排队叫号队列 Q, 部门表 P " & vbCrLf & _
                 " Where Q.科室ID=P.ID and Q.患者姓名 like [2]"
    End If

    Set FindQueueData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, str门诊号, str姓名, str就诊卡号, str医保号)
    
    Exit Function
errHandle:
    Set FindQueueData = Nothing
    If ErrCenter = 1 Then Resume

End Function



Private Sub LoadDataToFace(lvwData As ListView, rsData As ADODB.Recordset, strKey As String)
'**************************************************************************************************
'载入查询的数据信息到ListView中
'
'lvwQueueData：负责数据显示
'rsData：数据源
'strKey：保存关键字
'
'**************************************************************************************************
    
    On Error GoTo errHandle

    '清除所有数据
    Call lvwData.ListItems.Clear
    
    If rsData.RecordCount <= 0 Then Exit Sub
      
    Dim i As Integer
      
    Call rsData.MoveFirst
      
        
    '循环读取数据
    While Not rsData.EOF
      Dim liRow As ListItem
      
      Set liRow = lvwData.ListItems.Add()
      
      'liRow.SmallIcon = 1
      'liRow.Icon = 1
      
      '当使用RestoreWinState过程后，listview关键字前会自动添加"_"
      '读取第一列信息
      If Not IsNull(rsData.Fields.Item(Replace(lvwData.ColumnHeaders(1).Key, "_", ""))) Then
        liRow.Text = rsData.Fields.Item(Replace(lvwData.ColumnHeaders(1).Key, "_", ""))
      Else
        liRow.Text = ""
      End If
      
      '读取关键字
      liRow.Tag = rsData(strKey)
      
      For i = 2 To lvwData.ColumnHeaders.Count
        Dim liSubItem As ListSubItem
        
        Set liSubItem = liRow.ListSubItems.Add()
        
        If Not IsNull(rsData.Fields.Item(Replace(lvwData.ColumnHeaders(i).Key, "_", ""))) Then
          liSubItem.Text = rsData.Fields.Item(Replace(lvwData.ColumnHeaders(i).Key, "_", ""))
        Else
          liSubItem.Text = ""
        End If
    
      Next i
          
      Call rsData.MoveNext
    Wend
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub


Private Sub Execute_恢复(ByVal Id As Long)
    On Error GoTo errHandle
        
        Dim strSql As String
        
        strSql = "ZL_排队叫号队列_恢复(" & Id & ")"
                
        Call zlDatabase.ExecuteProcedure(strSql, "复诊")
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub


Private Sub Execute_回诊(ByVal Id As Long)
    On Error GoTo errHandle
        
        Dim strSql As String
        
        strSql = "ZL_排队叫号队列_回诊(" & Id & ")"
                
        Call zlDatabase.ExecuteProcedure(strSql, "复诊")
    Exit Sub
errHandle:
    If ErrCenter = 1 Then Resume
End Sub


Private Sub Form_Load()
    '恢复窗口状态
    Call RestoreWinState(Me, App.ProductName)

    cboFindWay.ListIndex = 1
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    
    lvwQueueData.Left = 100
    lvwQueueData.Top = Picture1.Height + 100
    lvwQueueData.Width = Me.Width - 200
    lvwQueueData.Height = Picture2.Top - Picture1.Height - 200
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub Picture2_Resize()
    On Error Resume Next
    
    cmdExit(0).Left = Me.Width - cmdExit(0).Width - 200
    cmdExit(2).Left = cmdExit(0).Left - cmdExit(2).Width - 50
    cmdExit(1).Left = cmdExit(2).Left - cmdExit(1).Width - 50
End Sub

Private Sub txtFindData_KeyPress(KeyAscii As Integer)
    Dim blnCard As Boolean
    Dim rsData As ADODB.Recordset
    
    If KeyAscii = 13 Then
        Call cmdStartFind_Click
        Exit Sub
    End If
    
    mstrFindKey = cboFindWay.Text
    
    If mstrFindKey = "门诊号" Then  '门诊号
        If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    ElseIf mstrFindKey = "就诊卡号" Or mstrFindKey = "姓名" Then     '就诊卡号,'姓名处也是刷卡
            If InStr(":：;；?？", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
            
            blnCard = zlCommFun.InputIsCard(txtFindData, KeyAscii, glngSys)
            If blnCard And Len(txtFindData.Text) = gbyt磁卡 - 1 And KeyAscii <> 8 Then
            
                txtFindData.Text = txtFindData.Text & Chr(KeyAscii)
                txtFindData.SelStart = Len(txtFindData.Text)
                
                KeyAscii = 0
                
                Call lvwQueueData.ListItems.Clear
                
                Set rsData = FindQueueData("就诊卡号", txtFindData.Text)
                
                If rsData Is Nothing Then
                    MsgBox "没有检索到所需数据。", vbInformation, "排队叫号系统"
                    Exit Sub
                End If
                
                If rsData.RecordCount <= 0 Then
                    MsgBox "没有检索到所需数据。", vbInformation, "排队叫号系统"
                    Exit Sub
                End If
                
                Call LoadDataToFace(lvwQueueData, rsData, "ID")
                
            End If
    ElseIf mstrFindKey = "医保号" Then    '医保号
    End If
End Sub
