VERSION 5.00
Begin VB.Form frmSelTablespace 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "��������"
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
   StartUpPosition =   1  '����������
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
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   6480
         TabIndex        =   12
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
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
      ToolTipText     =   "����ִ�пɴ����������ٶȣ�������������ʱ��Ȼ�Ὣ���ݷŵ��ļ�ĩβ��������������������ò��ж�Ϊ0"
      Top             =   3120
      Width           =   375
   End
   Begin VB.CheckBox chkOnline 
      Appearance      =   0  'Flat
      Caption         =   "������������"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      ToolTipText     =   "��ѡ���ٶȻ����½������Ҳ���������־�����ǲ���Ӱ�쵱ǰҵ�������ʹ��"
      Top             =   3120
      Width           =   1400
   End
   Begin VB.Frame fraLine 
      Caption         =   "����ģʽ"
      Height          =   2175
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   7455
      Begin VB.OptionButton optMode 
         Caption         =   "�ڵ�ǰ��ռ��������ٶȽϿ죬���пռ�Ҫ���С��������������ļ�ĩβ�������ݣ�"
         Height          =   200
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   420
         Value           =   -1  'True
         Width           =   7215
      End
      Begin VB.OptionButton optMode 
         Caption         =   "������������ռ䣬��ɺ��Լ�ȥɾ���ɵı�ռ䣨�ٶȽϿ죬���пռ�Ҫ�����"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   825
         Width           =   7095
      End
      Begin VB.OptionButton optMode 
         Caption         =   "�������ݴ��ռ䣬����ԭ��ռ��ļ������ƻ������ٶ����������п���Ҫ���С��"
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
         Caption         =   "��ռ�����"
         Height          =   255
         Left            =   600
         TabIndex        =   6
         Top             =   1680
         Width           =   1095
      End
   End
   Begin VB.Label lblSortSize 
      BackStyle       =   0  'Transparent
      Caption         =   "ÿ�����̵��������ڴ��С       M"
      Height          =   255
      Left            =   3600
      TabIndex        =   14
      Top             =   3150
      Width           =   3090
   End
   Begin VB.Label lblParallel 
      BackStyle       =   0  'Transparent
      Caption         =   "�������ж�"
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
      Caption         =   "����ݿռ��ʱ��Ĳ�ͬ�������ѡ���ʺϵĲ���"
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
Private mblnAdjusted As Boolean '�Ƿ��Ѿ��������Ự����

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
            strPrompt = "������ı�ռ�"
        Else
            strSQL = "Select 1 From DBA_TABLESPACES Where TABLESPACE_NAME = [1]"
            Set rsTbs = OpenSQLRecord(strSQL, "��ռ���", mstrTbs)
            If rsTbs.RecordCount = 0 Then strPrompt = "ָ���ı�ռ䲻���ڣ�����������"
        End If
        
        If strPrompt <> "" Then
            MsgBox strPrompt, vbExclamation, "��ʾ"
            If txtTBS.Enabled And txtTBS.Visible Then txtTBS.SetFocus
            Exit Sub
        End If
    End If
    
    If Val(txtSortSize.Text) > 2048 Then
        MsgBox "ÿ�����̵��������ڴ��С���ܳ���2G(2048M)", vbInformation, gstrSysName
        Exit Sub
    ElseIf Val(txtParallel.Text) > 0 Then
        If MsgBox("ע�⣺�������ݿ��������ʣ������ڴ��Ƿ���" & Val(txtParallel.Text) * Val(txtSortSize.Text) & "M,����ڴ治�㣬���ܵ��²���ʧ�ܻ�������������ڴ�ľ�������Ӧ��", vbOKCancel + vbDefaultButton1, "����") = vbCancel Then
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
        
        'ֱ��·��IO�Ĵ�С
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
        gcnOracle.Execute strSQL '����10G��BUG����Ҫִ�����β���Ч
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
'���ܣ���ȡ����ʾ���ж�
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errH
    strSQL = "Select Value From V$parameter Where Name = 'cpu_count'"
    Set rsTmp = OpenSQLRecord(strSQL, Me.Caption)
    If rsTmp.EOF Then
        txtParallel.Text = "0"
        txtParallel.Locked = True
        txtParallel.Enabled = False
        lblParallel.ToolTipText = "δ�ܶ�ȡ�����ݿ����cpu_count"
    Else
        txtParallel.Tag = "" & rsTmp!Value
        If Val(rsTmp!Value) < 3 Then
            txtParallel.Text = "0"
            txtParallel.Enabled = False
            lblParallel.ToolTipText = "������Cpu��������3�������ܽ��в���ִ��"
        ElseIf Val(rsTmp!Value) < 13 Then
            txtParallel.Text = Val(rsTmp!Value) \ 2 'һ��ȡ��
        Else
            txtParallel.Text = "12"  '��ʹcpu�㹻�����Կ��������ڴ������ܣ����жȲ���Խ��Խ��
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
            MsgBox "���жȲ��ܳ���cpu����" & txtParallel.Tag, vbInformation, gstrSysName
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
            MsgBox "ÿ�����̵��������ڴ��С���ܳ���2G(2048M)", vbInformation, gstrSysName
            Cancel = True
        End If
    End If
End Sub
