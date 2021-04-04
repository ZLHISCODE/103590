VERSION 5.00
Begin VB.Form frmSelTablespace 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "重整参数"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   7710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtSortSize 
      Alignment       =   1  'Right Justify
      Height          =   280
      Left            =   5790
      MaxLength       =   4
      TabIndex        =   13
      Text            =   "2048"
      ToolTipText     =   $"frmSelTablespace.frx":0000
      Top             =   3120
      Width           =   495
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00EFF0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   0
      ScaleHeight     =   645
      ScaleWidth      =   7710
      TabIndex        =   10
      Top             =   3525
      Width           =   7710
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   6480
         TabIndex        =   12
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Default         =   -1  'True
         Height          =   350
         Left            =   5280
         TabIndex        =   11
         Top             =   120
         Width           =   1100
      End
      Begin VB.Line lineBottom 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   0
         X2              =   10000
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line lineBottom 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   0
         X2              =   10000
         Y1              =   15
         Y2              =   15
      End
   End
   Begin VB.TextBox txtParallel 
      Alignment       =   1  'Right Justify
      Height          =   280
      Left            =   2730
      TabIndex        =   8
      Text            =   "12"
      ToolTipText     =   "并行执行可大幅提高重整速度，但是重整后有时仍然会将数据放到文件末尾，遇到这种情况，请设置并行度为0"
      Top             =   3120
      Width           =   375
   End
   Begin VB.CheckBox chkOnline 
      Appearance      =   0  'Flat
      Caption         =   "在线重整索引"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      ToolTipText     =   "勾选后速度会大幅下降，并且产生大量日志，但是不会影响当前业务的正常使用"
      Top             =   3120
      Width           =   1400
   End
   Begin VB.Frame fraLine 
      Caption         =   "重整模式"
      Height          =   2175
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   7455
      Begin VB.OptionButton optMode 
         Caption         =   "在当前表空间重整（速度较快，空闲空间要求较小，但重整后可能文件末尾还有数据）"
         Height          =   200
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   420
         Value           =   -1  'True
         Width           =   7215
      End
      Begin VB.OptionButton optMode 
         Caption         =   "重整到其他表空间，完成后自己去删除旧的表空间（速度较快，空闲空间要求最大）"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   825
         Width           =   7095
      End
      Begin VB.OptionButton optMode 
         Caption         =   "重整到暂存表空间，收缩原表空间文件后再移回来（速度最慢，空闲空闲要求较小）"
         Height          =   180
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   6975
      End
      Begin VB.TextBox txtTBS 
         Height          =   300
         Left            =   1680
         TabIndex        =   2
         Text            =   "SYSAUX"
         Top             =   1620
         Width           =   1455
      End
      Begin VB.Label lblTbs 
         Caption         =   "表空间名称"
         Height          =   255
         Left            =   600
         TabIndex        =   6
         Top             =   1680
         Width           =   1095
      End
   End
   Begin VB.Label lblSortSize 
      BackStyle       =   0  'Transparent
      Caption         =   "每个进程的排序区内存大小       M"
      Height          =   255
      Left            =   3600
      TabIndex        =   14
      Top             =   3150
      Width           =   3090
   End
   Begin VB.Label lblParallel 
      BackStyle       =   0  'Transparent
      Caption         =   "重整并行度"
      Height          =   255
      Left            =   1800
      TabIndex        =   9
      Top             =   3150
      Width           =   930
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmSelTablespace.frx":008B
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblPrompt 
      Caption         =   "请根据空间和时间的不同需求情况选择适合的参数"
      Height          =   255
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   3975
   End
End
Attribute VB_Name = "frmSelTablespace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbytMode As Byte
Private mstrTbs As String
Private mblnOK As Boolean
Private mstrParallel As String
Private mstrOnline As String
Private mblnAdjusted As Boolean '是否已经调整过会话参数

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strSQL As String, strPrompt As String
    Dim rsTbs As ADODB.Recordset
    Dim i As Double
    
    On Error GoTo errH
    
    If optMode(0).Value = False Then
        mstrTbs = UCase(Trim(txtTBS.Text))
        If mstrTbs = "" Then
            strPrompt = "请输入的表空间"
        Else
            strSQL = "Select 1 From DBA_TABLESPACES Where TABLESPACE_NAME = [1]"
            Set rsTbs = OpenSQLRecord(strSQL, "表空间检查", mstrTbs)
            If rsTbs.RecordCount = 0 Then strPrompt = "指定的表空间不存在，请重新输入"
        End If
        
        If strPrompt <> "" Then
            MsgBox strPrompt, vbExclamation, "提示"
            If txtTBS.Enabled And txtTBS.Visible Then txtTBS.SetFocus
            Exit Sub
        End If
    End If
    
    If Val(txtSortSize.Text) > 2048 Then
        MsgBox "每个进程的排序区内存大小不能超过2G(2048M)", vbInformation, gstrSysName
        Exit Sub
    ElseIf Val(txtParallel.Text) > 0 Then
        If MsgBox("注意：请检查数据库服务器的剩余空闲内存是否有" & Val(txtParallel.Text) * Val(txtSortSize.Text) & "M,如果内存不足，可能导致操作失败或操作过程中因内存耗尽而无响应。", vbOKCancel + vbDefaultButton1, "提醒") = vbCancel Then
            Exit Sub
        End If
    End If
    
    
    For i = 0 To optMode.Count - 1
        If optMode(i).Value Then mbytMode = i: Exit For
    Next
        
    If txtParallel.Text <> "0" Then mstrParallel = " Parallel " & txtParallel.Text
    If chkOnline.Value = 1 Then mstrOnline = "Online"
    
    If mblnAdjusted = False Then
        mblnAdjusted = True
        strSQL = "alter session set workarea_size_policy=MANUAL"
        gcnOracle.Execute strSQL
        
        '直接路径IO的大小
        strSQL = "alter session set events '10351 trace name context forever, level 128'"
        gcnOracle.Execute strSQL
        
        strSQL = "alter session SET db_file_multiblock_read_count=128"
        gcnOracle.Execute strSQL
        
        strSQL = "alter session set ""_sort_multiblock_read_count""=128"
        gcnOracle.Execute strSQL
                
        strSQL = "alter session SET db_block_checking=false"
        gcnOracle.Execute strSQL
    End If
    
    If txtSortSize.Text <> "0" Then
        If txtSortSize.Text = "2048" Then
            i = CDbl(txtSortSize.Text) * 1024 * 1024 - 1
        Else
            i = CDbl(txtSortSize.Text) * 1024 * 1024
        End If
        
        strSQL = "alter session SET sort_area_size=" & i
        gcnOracle.Execute strSQL
        gcnOracle.Execute strSQL '由于10G的BUG，需要执行两次才生效
    Else
        strSQL = "alter session set workarea_size_policy=auto"
        gcnOracle.Execute strSQL
    End If
    
    mblnOK = True
    Unload Me
    
    Exit Sub
errH:
    Call ErrCenter(strSQL)
End Sub

Public Function ShowMe(frmParent As Form, bytMode As Byte, strTbs As String, strParallel As String, strOnline As String) As Boolean
    mstrTbs = ""
    mstrParallel = ""
    mstrOnline = ""
    
    Me.Show vbModal, frmParent
    
    bytMode = mbytMode
    strTbs = mstrTbs
    strParallel = mstrParallel
    strOnline = mstrOnline
    
    ShowMe = mblnOK
End Function


Private Sub LoadParallel()
'功能：读取并显示并行度
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errH
    strSQL = "Select Value From V$parameter Where Name = 'cpu_count'"
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption)
    If rsTmp.EOF Then
        txtParallel.Text = "0"
        txtParallel.Locked = True
        txtParallel.Enabled = False
        lblParallel.ToolTipText = "未能读取到数据库参数cpu_count"
    Else
        txtParallel.Tag = "" & rsTmp!Value
        If Val(rsTmp!Value) < 3 Then
            txtParallel.Text = "0"
            txtParallel.Enabled = False
            lblParallel.ToolTipText = "服务器Cpu数量不足3个，不能进行并行执行"
        ElseIf Val(rsTmp!Value) < 13 Then
            txtParallel.Text = Val(rsTmp!Value) \ 2 '一半取整
        Else
            txtParallel.Text = "12"  '即使cpu足够，但仍可能受限于磁盘性能，并行度并非越大越好
        End If
    End If

    Exit Sub
errH:
    Call ErrCenter(strSQL)
End Sub

Private Sub Form_Load()
    
    Call LoadParallel
    
    Call optMode_Click(0)
End Sub

Private Sub optMode_Click(Index As Integer)
       
    txtTBS.Enabled = Index <> 0
    If Index = 1 Then
        txtTBS.Text = ""
        txtTBS.SetFocus
    ElseIf Index = 2 Then
        txtTBS.Text = "SYSAUX"
        txtTBS.SetFocus
    End If
End Sub

Private Sub txtTBS_GotFocus()
    If txtTBS.Text <> "" Then
        txtTBS.SelStart = 0
        txtTBS.SelLength = Len(txtTBS.Text)
    End If
End Sub

Private Sub txtTBS_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        cmdOK.SetFocus
    End If
End Sub


Private Sub txtParallel_GotFocus()
    txtParallel.SelStart = 0
    txtParallel.SelLength = Len(txtParallel.Text)
End Sub

Private Sub txtParallel_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtParallel_Validate(Cancel As Boolean)
    If Val(txtParallel.Tag) <> 0 Then
        If Val(txtParallel.Text) > Val(txtParallel.Tag) Then
            MsgBox "并行度不能超过cpu个数" & txtParallel.Tag, vbInformation, gstrSysName
            Cancel = True
        End If
    End If
End Sub

Private Sub txtSortSize_GotFocus()
    txtSortSize.SelStart = 0
    txtSortSize.SelLength = Len(txtSortSize.Text)
End Sub

Private Sub txtSortSize_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtSortSize_Validate(Cancel As Boolean)
    If Val(txtSortSize.Tag) <> 0 Then
        
        If Val(txtSortSize.Text) > 2048 Then
            MsgBox "每个进程的排序区内存大小不能超过2G(2048M)", vbInformation, gstrSysName
            Cancel = True
        End If
    End If
End Sub
