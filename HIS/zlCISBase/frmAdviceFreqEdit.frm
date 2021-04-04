VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAdviceFreqEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "医嘱频率编辑"
   ClientHeight    =   3015
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4830
   Icon            =   "frmAdviceFreqEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin MSComCtl2.UpDown UD频率间隔 
      Height          =   300
      Left            =   2835
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   2145
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "txtEdit(5)"
      BuddyDispid     =   196611
      BuddyIndex      =   5
      OrigLeft        =   2836
      OrigTop         =   2145
      OrigRight       =   3076
      OrigBottom      =   2445
      Max             =   99
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown UD频率次数 
      Height          =   300
      Left            =   2835
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1740
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "txtEdit(4)"
      BuddyDispid     =   196611
      BuddyIndex      =   4
      OrigLeft        =   2836
      OrigTop         =   1743
      OrigRight       =   3076
      OrigBottom      =   2043
      Max             =   99
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Index           =   3
      Left            =   1290
      MaxLength       =   50
      TabIndex        =   2
      Top             =   1341
      Width           =   1785
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   1290
      MaxLength       =   10
      TabIndex        =   1
      Top             =   939
      Width           =   1785
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   3570
      TabIndex        =   8
      Top             =   2295
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   3510
      Left            =   3345
      TabIndex        =   15
      Top             =   -300
      Width           =   30
   End
   Begin VB.ComboBox cbo间隔单位 
      Height          =   300
      IMEMode         =   3  'DISABLE
      ItemData        =   "frmAdviceFreqEdit.frx":000C
      Left            =   1290
      List            =   "frmAdviceFreqEdit.frx":001C
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2550
      Width           =   1785
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Index           =   5
      Left            =   1290
      MaxLength       =   2
      TabIndex        =   4
      Text            =   "1"
      Top             =   2145
      Width           =   1530
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Index           =   4
      Left            =   1290
      MaxLength       =   2
      TabIndex        =   3
      Text            =   "1"
      Top             =   1743
      Width           =   1530
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      Index           =   1
      Left            =   1290
      MaxLength       =   20
      TabIndex        =   0
      Top             =   537
      Width           =   1785
   End
   Begin VB.TextBox txtEdit 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Index           =   0
      Left            =   1290
      MaxLength       =   3
      TabIndex        =   10
      Top             =   135
      Width           =   1785
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3570
      TabIndex        =   7
      Top             =   840
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3570
      TabIndex        =   6
      Top             =   390
      Width           =   1100
   End
   Begin VB.Label lbl英文名称 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "英文名称(&E)"
      Height          =   180
      Left            =   225
      TabIndex        =   17
      Top             =   1395
      Width           =   990
   End
   Begin VB.Label lbl简码 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "简码(&S)"
      Height          =   180
      Left            =   585
      TabIndex        =   16
      Top             =   1005
      Width           =   630
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "间隔单位(&U)"
      Height          =   180
      Left            =   225
      TabIndex        =   14
      Top             =   2610
      Width           =   990
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "频率间隔(&J)"
      Height          =   180
      Left            =   225
      TabIndex        =   13
      Top             =   2205
      Width           =   990
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "频率次数(&M)"
      Height          =   180
      Left            =   225
      TabIndex        =   12
      Top             =   1800
      Width           =   990
   End
   Begin VB.Label lbl名称 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "名称(&N)"
      Height          =   180
      Left            =   585
      TabIndex        =   11
      Top             =   600
      Width           =   630
   End
   Begin VB.Label lbl编码 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "编码(&B)"
      Height          =   180
      Left            =   585
      TabIndex        =   9
      Top             =   195
      Width           =   630
   End
End
Attribute VB_Name = "frmAdviceFreqEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mstrCode As String
Public mbytType As Byte '1-西医,2-中医

Private mstr间隔单位 As String
Private mint频率次数 As Integer
Private mint频率间隔 As Integer

Private mblnChange As Boolean

Private Sub cbo间隔单位_Click()
    If Visible Then mblnChange = True
    
    If cbo间隔单位.Text = "周" Then
        txtEdit(5).Enabled = False
        UD频率间隔.Enabled = False
        txtEdit(5).Text = 1
    ElseIf cbo间隔单位.Text = "分钟" Then
        txtEdit(4).Enabled = False
        UD频率次数.Enabled = False
        txtEdit(4).Text = 1
    Else
        txtEdit(5).Enabled = True
        UD频率间隔.Enabled = True
        txtEdit(4).Enabled = True
        UD频率次数.Enabled = True
    End If
    
    If cbo间隔单位.Text = "分钟" Then
        txtEdit(5).MaxLength = 3
        UD频率间隔.Max = 999
    Else
        txtEdit(5).MaxLength = 2
        UD频率间隔.Max = 99
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    Dim strDel As String
    Dim blnDel As Boolean
    
    If txtEdit(0).Text = "" Then
        MsgBox "必须输入编码。", vbInformation, gstrSysName
        txtEdit(0).SetFocus: Exit Sub
    End If
    
    If txtEdit(1).Text = "" Then
        MsgBox "必须输入名称。", vbInformation, gstrSysName
        txtEdit(1).SetFocus: Exit Sub
    End If
    
    strSql = "Select 1 From 诊疗频率项目 Where  Nvl(适用范围, 0) <> 1 And Nvl(适用范围, 0) <> 2 And 名称=[1] and rownum<2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "cmdOK_Click", txtEdit(1).Text)
    
    If rsTmp.RecordCount > 0 Then
        MsgBox "固定频率中已经有了相同名称的频率。", vbInformation, gstrSysName
        txtEdit(1).SetFocus: Exit Sub
    End If
    
    If zlCommFun.ActualLen(txtEdit(1).Text) > txtEdit(1).MaxLength Then
        MsgBox "名称内容过长，最多只能包含" & txtEdit(1).MaxLength & "个字符或" & txtEdit(1).MaxLength \ 2 & "个汉字。", vbInformation, gstrSysName
        txtEdit(1).SetFocus: Exit Sub
    End If
    
'    If Val(txtEdit(4).Text) <> 1 And Val(txtEdit(5).Text) <> 1 Then
'        MsgBox "频率次数和频率间隔两者应该有一个为 1 。", vbInformation, gstrSysName
'        txtEdit(5).SetFocus: Exit Sub
'    End If
        
    If mstrCode = "" Then
        strSql = "ZL_诊疗频率项目_Insert('" & txtEdit(0).Text & "','" & txtEdit(1).Text & "','" & txtEdit(2).Text & "','" & txtEdit(3).Text & "'," & txtEdit(4).Text & "," & txtEdit(5).Text & ",'" & cbo间隔单位.Text & "'," & mbytType & ")"
    Else
        If Val(txtEdit(4).Text) <> mint频率次数 Then
            If MsgBox("你更改了频率次数，这将清除该频率项目现有的时间设置。要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            blnDel = True
        ElseIf Val(txtEdit(5).Text) <> mint频率间隔 And cbo间隔单位.Text <> "分钟" Then
            If MsgBox("你更改了频率间隔，这将清除该频率项目现有的时间设置。要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            blnDel = True
        ElseIf cbo间隔单位.Text <> mstr间隔单位 Then
            If MsgBox("你更改了间隔单位，这将清除该频率项目现有的时间设置。要继续吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            blnDel = True
        End If
        If blnDel Then strDel = "ZL_诊疗频率时间_Delete('" & mstrCode & "')"
        strSql = "ZL_诊疗频率项目_Update('" & mstrCode & "','" & txtEdit(0).Text & "','" & txtEdit(1).Text & "','" & txtEdit(2).Text & "','" & txtEdit(3).Text & "'," & txtEdit(4).Text & "," & txtEdit(5).Text & ",'" & cbo间隔单位.Text & "')"
    End If
        
    On Error GoTo errH
    gcnOracle.BeginTrans
    If strDel <> "" Then
        Call zlDatabase.ExecuteProcedure(strDel, Me.Caption)
    End If
    Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
    gcnOracle.CommitTrans
    On Error GoTo 0
    
    If mstrCode <> "" Then
        mblnChange = False
        gblnOK = True
        Unload Me
    Else
        Call frmAdviceFreq.LoadItems("_" & txtEdit(0).Text)
        Call Form_Load
        mblnChange = False
        gblnOK = True
        txtEdit(1).SetFocus
    End If
    Exit Sub
errH:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf Chr(KeyAscii) = "'" Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    
    gblnOK = False
    mblnChange = False
    
    On Error GoTo errH
    
    If mstrCode <> "" Then
        '修改
        strSql = "Select * From 诊疗频率项目 Where 编码=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mstrCode)
                
        cbo间隔单位.Text = Nvl(rsTmp!间隔单位)
        txtEdit(0).Text = Nvl(rsTmp!编码)
        txtEdit(1).Text = Nvl(rsTmp!名称)
        txtEdit(2).Text = Nvl(rsTmp!简码)
        txtEdit(3).Text = Nvl(rsTmp!英文名称)
        txtEdit(4).Text = Nvl(rsTmp!频率次数)
        txtEdit(5).Text = Nvl(rsTmp!频率间隔)
        
        mstr间隔单位 = Nvl(rsTmp!间隔单位)
        mint频率次数 = Nvl(rsTmp!频率次数)
        mint频率间隔 = Nvl(rsTmp!频率间隔)
    Else
        '新增
        txtEdit(0).Text = ""
        txtEdit(1).Text = ""
        txtEdit(2).Text = ""
        txtEdit(3).Text = ""
        
        strSql = "Select ZL_IncStr(Max(编码)) as 编码 From 诊疗频率项目"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
        If Not rsTmp.EOF Then
            If Not IsNull(rsTmp!编码) Then
                txtEdit(0).Text = rsTmp!编码
            End If
        End If
        If txtEdit(0).Text = "" Then
            txtEdit(0).Text = String(txtEdit(0).MaxLength - 1, "0") & "1"
        End If
        If cbo间隔单位.ListIndex = -1 Then cbo间隔单位.Text = "天"
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange Then
        If MsgBox("你已修改了相关内容，确实要退出吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = True: Exit Sub
        End If
    End If
    
    mstrCode = ""
    mbytType = 0
End Sub

Private Sub txtEdit_Change(Index As Integer)
    If Visible Then mblnChange = True
    '产生简码
    If Index = 1 And Visible Then txtEdit(2).Text = zlCommFun.SpellCode(txtEdit(Index).Text)
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
    If txtEdit(Index).IMEMode = 0 Then Call zlCommFun.OpenIme(True)
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
    If txtEdit(Index).IMEMode = 0 Then Call zlCommFun.OpenIme
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 0 Then
        If InStr("-", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    ElseIf Index = 4 Or Index = 5 Then
        If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txtEdit_Validate(Index As Integer, Cancel As Boolean)
    If Not zlCommFun.StrIsValid(txtEdit(Index).Text, txtEdit(Index).MaxLength) Then
        Cancel = True
    ElseIf Index = 4 Or Index = 5 Then
        If Not IsNumeric(txtEdit(Index).Text) Or Val(txtEdit(Index).Text) <= 0 Then
            Cancel = True
            MsgBox "必须是数字且大于零！"
        End If
    End If
End Sub
