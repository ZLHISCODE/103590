VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmStPathItemEdit 
   Caption         =   "段落新增"
   ClientHeight    =   6075
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5910
   Icon            =   "frmStPathItemEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   5910
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      BackColor       =   &H00F0F4E4&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   5910
      TabIndex        =   7
      Top             =   5460
      Width           =   5910
      Begin VB.CommandButton cmdCancel 
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   4680
         TabIndex        =   9
         Top             =   160
         Width           =   1100
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "确定(&O)"
         Height          =   350
         Left            =   3240
         TabIndex        =   8
         Top             =   160
         Width           =   1100
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   0
         X2              =   10000
         Y1              =   45
         Y2              =   45
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   0
         X2              =   10000
         Y1              =   30
         Y2              =   30
      End
   End
   Begin VB.ComboBox cboNo 
      Height          =   300
      Left            =   840
      TabIndex        =   1
      Text            =   "cboNO"
      Top             =   82
      Width           =   2295
   End
   Begin VB.CheckBox chkContinual 
      Caption         =   "连续插入"
      Height          =   225
      Left            =   4675
      TabIndex        =   6
      Top             =   120
      Value           =   1  'Checked
      Width           =   1100
   End
   Begin RichTextLib.RichTextBox rtfContent 
      Height          =   4335
      Left            =   840
      TabIndex        =   5
      Top             =   1080
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   7646
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmStPathItemEdit.frx":59D62
   End
   Begin VB.TextBox txtTitle 
      Height          =   300
      Left            =   840
      TabIndex        =   3
      Top             =   581
      Width           =   4935
   End
   Begin VB.Label lblContent 
      Caption         =   "内容(&Q)"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   720
   End
   Begin VB.Label lblItemTile 
      Caption         =   "标题(&T)"
      Height          =   225
      Left            =   120
      TabIndex        =   2
      Top             =   619
      Width           =   720
   End
   Begin VB.Label lblNO 
      Caption         =   "序号(&N)"
      Height          =   225
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   720
   End
End
Attribute VB_Name = "frmStPathItemEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mintMode As Integer '0-增加流程项目，1-修改路径流程项目，2-删除路径流程项目
Private mlngStPathID As Long
Private mlng序号 As Long
Private mrsCourseItems As New ADODB.Recordset
Private mblnOK As Boolean '是否进行了数据操作

Public Function ShowMe(ByRef FrmParent As Object, ByVal intMode As Integer, ByVal lngStPathID As Long, Optional ByVal lng序号 As Long) As Boolean
'功能：显示路径流程项目增删改界面
'参数： intMode 0-增加流程项目，1-修改路径流程项目，2-删除路径流程项目
'       lngStPathID 标准路径ID
'       lng序号 路径流程项目序号

    mintMode = intMode
    mlngStPathID = lngStPathID
    mlng序号 = lng序号
    mblnOK = False
    Me.Show 1, FrmParent
    ShowMe = mblnOK
    
End Function


Private Sub cboNo_Click()
'功能：点击序号下拉列表，更新界面数据
    Dim strSel As String
    
    If cboNo.ListIndex = -1 Then Exit Sub
    
    strSel = cboNo.List(cboNo.ListIndex)
    
    If InStr(strSel, "-") > 0 Then
        cboNo.Text = Mid(strSel, 1, InStr(strSel, "-") - 1)
    Else
        cboNo.Text = strSel
    End If
    mlng序号 = Val(cboNo.Text)
    
    If Me.Visible Then
        mrsCourseItems.Filter = "序号=" & Val(cboNo.Text)
        If mrsCourseItems.RecordCount <> 0 Then
            txtTitle.Text = IIf(mintMode <> 0, mrsCourseItems!标题 & "", "")
            rtfContent.Text = IIf(mintMode <> 0, mrsCourseItems!内容 & "", "")
        End If
        mrsCourseItems.Filter = ""
    Else '初始加载下拉列表
        txtTitle.Text = IIf(mintMode <> 0, mrsCourseItems!标题 & "", "")
        rtfContent.Text = IIf(mintMode <> 0, mrsCourseItems!内容 & "", "")
    End If
    
End Sub

Private Sub cboNo_KeyPress(KeyAscii As Integer)
'功能：输入检测

    '只允许输入数字以及回车
    If Not (InStr("0123456789", Chr(KeyAscii)) > 0 Or KeyAscii = vbKeyReturn Or KeyAscii = vbKeyBack) Then KeyAscii = 0: Exit Sub
    'if KeyAscii = vbKeyReturn then
    '在非插入路径流程项目的情况下，输入数字不允许大于最大序号
    If Val(cboNo.Text & Chr(KeyAscii)) > mrsCourseItems.RecordCount And mintMode <> 0 Then KeyAscii = 0: Exit Sub
    '在插入路径流程项目的情况下，输入数字不允许大于最大序号+1
    If Val(cboNo.Text & Chr(KeyAscii)) > mrsCourseItems.RecordCount + 1 And mintMode = 0 Then KeyAscii = 0: Exit Sub

End Sub

Private Sub cboNo_LostFocus()
    Call cboNo_Click
End Sub

Private Sub cboNo_Validate(Cancel As Boolean)
    Dim strSel As String
    
    strSel = cboNo.Text
    
    If InStr(strSel, "-") > 0 Then
        cboNo.Text = Mid(strSel, 1, InStr(strSel, "-") - 1)
    Else
        cboNo.Text = strSel
    End If
    
    If Val(cboNo.Text) = 0 Then
        cboNo.Text = mlng序号
    Else
        mlng序号 = Val(cboNo.Text)
    End If
    
     '在非插入路径流程项目的情况下，输入数字不允许大于最大序号
    If Val(cboNo.Text) > mrsCourseItems.RecordCount And mintMode <> 0 Then
        mlng序号 = mrsCourseItems.RecordCount
        cboNo.Text = mlng序号
        Exit Sub
    End If
    
    '在插入路径流程项目的情况下，输入数字不允许大于最大序号+1
    If Val(cboNo.Text) > mrsCourseItems.RecordCount + 1 And mintMode = 0 Then
        mlng序号 = mrsCourseItems.RecordCount + 1
        cboNo.Text = mlng序号
        Exit Sub
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdUpdate_Click()
'功能：进行数据跟新，并且根据是否连续操作，来初始化下次界面或退出
    Dim strSql As String
    
    mblnOK = True
    
    On Error GoTo errH
    Select Case mintMode
        Case 0
            strSql = "Zl_标准路径流程_Insert(" & mlngStPathID & "," & mlng序号 & ",'" & Trim(txtTitle.Text) & "','" & Trim(rtfContent.Text) & "')"
        Case 1
            strSql = "Zl_标准路径流程_Update(" & mlngStPathID & "," & mlng序号 & ",'" & Trim(txtTitle.Text) & "','" & Trim(rtfContent.Text) & "')"
        Case 2
            strSql = "Zl_标准路径流程_Delete(" & mlngStPathID & "," & mlng序号 & ")"
    End Select
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    
    If chkContinual.Value = 0 Then
        Unload Me
    Else '连续操作-初始化界面数据

        Call GetCourseItems
        '删除完自动退出
        If mrsCourseItems.RecordCount = 0 And mintMode = 2 Then Unload Me
        '序号自增
        If mlng序号 <= mrsCourseItems.RecordCount Then
            mlng序号 = mlng序号 + IIf(mintMode <> 2, 1, 0)
        Else
            mlng序号 = mrsCourseItems.RecordCount + IIf(mintMode <> 2, 1, 0)
        End If
        '修改到最后一个时退出
        If mlng序号 > mrsCourseItems.RecordCount And mintMode = 1 Then Unload Me
        cboNo.Text = mlng序号
        
        If mintMode = 0 Then '插入时清空数据
            txtTitle.Text = ""
            rtfContent.Text = ""
        End If
        
        Call InitcboNo '加载下拉列表
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdUpdate_KeyPress(KeyAscii As Integer)
'功能：回车触发click事件
    If KeyAscii = vbKeyReturn Then
        Call cmdUpdate_Click
    End If
End Sub

Private Sub Form_Activate()

    If mintMode <> 2 Then
        txtTitle.SetFocus
        txtTitle.SelStart = 0
        txtTitle.SelLength = Len(txtTitle.Text)
    Else
        cboNo.SetFocus
    End If
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
'功能：回车定位下一个控件
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub Form_Load()
'功能：根据操作类型初始化界面
    Select Case mintMode
        Case 0
            chkContinual.Caption = "连续增加"
            Me.Caption = "增加段落"
        Case 1
            chkContinual.Caption = "连续修改"
            Me.Caption = "修改段落"
        Case 2
            chkContinual.Caption = "连续删除"
            Me.Caption = "删除段落"
            txtTitle.Locked = True
            rtfContent.Locked = True
    End Select

    Call InitcboNo
     
End Sub

Private Sub GetCourseItems()
'功能：获取当前标准路径的所有路径流程项目
    Dim strSql As String
    
    On Error GoTo errH
    strSql = "Select a.序号, a.标题, a.内容 From 标准路径流程 A Where 标准路径id = [1] Order By a.序号"
    Set mrsCourseItems = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlngStPathID)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub InitcboNo()
'功能：初始化下拉列表
    Dim i As Long
    
    Call GetCourseItems
    Call cboNo.Clear
    With mrsCourseItems
        If .RecordCount = 0 Then cboNo.Text = mlng序号: Exit Sub
        .MoveFirst
        For i = 1 To .RecordCount
            cboNo.AddItem !序号 & "-" & !标题
            
            If !序号 = mlng序号 Then
                cboNo.ListIndex = cboNo.NewIndex
            End If
            .MoveNext
        Next
        cboNo.Text = mlng序号
        If mintMode = 0 Then
            cboNo.AddItem .RecordCount + 1
        End If
    End With

End Sub

Private Sub Form_Resize()

    If Me.WindowState = vbMaximized Or Me.WindowState = vbMinimized Then Exit Sub
    '不允许改变窗体大小
    If Me.Width < 6150 Then Me.Width = 6150
    If Me.Height < 6650 Then Me.Height = 6650
    
End Sub

Private Sub rtfContent_GotFocus()
'功能：内容输入框获得焦点时内容全选
    rtfContent.SelStart = Len(rtfContent.Text)
End Sub

Private Sub txtTitle_GotFocus()
'功能：标题输入框获得焦点时内容全选
    txtTitle.SelStart = 0
    txtTitle.SelLength = Len(txtTitle.Text)
End Sub
