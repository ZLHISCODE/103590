VERSION 5.00
Begin VB.Form frm子医保身份验证 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "子医保身份验证"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3825
   Icon            =   "frm子医保身份验证.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   3825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmd取消 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   2490
      TabIndex        =   12
      Top             =   2610
      Width           =   1100
   End
   Begin VB.CommandButton cmd确定 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   1230
      TabIndex        =   11
      Top             =   2610
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   30
      TabIndex        =   10
      Top             =   2400
      Width           =   3735
   End
   Begin VB.TextBox txt帐户余额 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1140
      TabIndex        =   9
      Top             =   1860
      Width           =   2415
   End
   Begin VB.TextBox txt姓名 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1140
      TabIndex        =   7
      Top             =   1470
      Width           =   2415
   End
   Begin VB.TextBox txt医保号 
      Height          =   300
      Left            =   1140
      TabIndex        =   5
      Top             =   1080
      Width           =   2415
   End
   Begin VB.ComboBox cbo优惠类别 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1140
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   690
      Width           =   2415
   End
   Begin VB.ComboBox cbo子医保 
      Height          =   300
      Left            =   1140
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   300
      Width           =   2415
   End
   Begin VB.Label lbl帐户余额 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "帐户余额"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   360
      TabIndex        =   8
      Top             =   1920
      Width           =   720
   End
   Begin VB.Label lbl姓名 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "姓名"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   720
      TabIndex        =   6
      Top             =   1530
      Width           =   360
   End
   Begin VB.Label lbl医保号 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "医保号"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   540
      TabIndex        =   4
      Top             =   1140
      Width           =   540
   End
   Begin VB.Label lbl优惠类别 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "优惠类别"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   360
      TabIndex        =   2
      Top             =   750
      Width           =   720
   End
   Begin VB.Label lbl子医保 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "子医保"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   540
      TabIndex        =   0
      Top             =   360
      Width           =   540
   End
End
Attribute VB_Name = "frm子医保身份验证"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim marrData
Dim mstrReg As String
Dim mstr序号 As String
Private mstr姓名 As String
Dim i As Integer, j As Integer
Dim rsTemp As New ADODB.Recordset
Private mstrReturn As String             '保险序号|优惠类别|医保号|余额|停用

Public Function ShowME(ByVal STR姓名 As String) As String
    mstr姓名 = STR姓名
    mstrReturn = ""
    Me.Show 1
    ShowME = mstrReturn
End Function

Private Sub cbo子医保_Click()
    For i = 0 To j
        If Split(marrData(i), ";")(0) = cbo子医保.ItemData(cbo子医保.ListIndex) Then
            Call zlControl.CboLocate(cbo优惠类别, Split(marrData(i), ";")(1), False)
            Exit For
        End If
    Next
End Sub

Private Sub cmd取消_Click()
    Unload Me
End Sub

Private Sub cmd确定_Click()
    If Trim(txt姓名.Text) = "" Then
        MsgBox "请输入医保号按回车确定病人身份！", vbInformation, gstrSysName
        txt医保号.SetFocus
        Exit Sub
    End If
    If Me.txt姓名.Text <> mstr姓名 Then
        MsgBox "两个接口返回的病人姓名不同，请检查！", vbInformation, gstrSysName
        txt医保号.SetFocus
        Exit Sub
    End If
    
    mstrReturn = Me.cbo子医保.ItemData(Me.cbo子医保.ListIndex) & "|" & Me.cbo优惠类别.Text & "|" & Me.txt医保号.Text & "|" & txt帐户余额.Text & "|" & Val(txt帐户余额.Tag)
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_Load()
    With Me.cbo优惠类别
        .Clear
        .AddItem "普通"
        .AddItem "惠民"
        .AddItem "慈善"
        .AddItem "零差率"
        .ListIndex = 0
    End With
    
    Me.cbo子医保.Clear
    mstrReg = GetSetting("ZLSOFT", "公共全局", "下属医保接口", "")
    marrData = Split(mstrReg, ",")
    j = UBound(marrData)
    For i = 0 To j
        mstr序号 = mstr序号 & "," & Split(marrData(i), ";")(0)
    Next
    mstr序号 = Mid(mstr序号, 2)
    
    gstrSQL = " Select 序号,名称 From 保险类别 Where 序号 IN ([1])"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取子医保", mstr序号)
    
    For i = 0 To j
        rsTemp.Filter = "序号=" & Split(marrData(i), ";")(0)
        Me.cbo子医保.AddItem rsTemp!序号 & "-" & rsTemp!名称
        Me.cbo子医保.ItemData(Me.cbo子医保.NewIndex) = rsTemp!序号
    Next
    Me.cbo子医保.ListIndex = 0
End Sub

Private Sub txt医保号_GotFocus()
    Call zlControl.TxtSelAll(txt医保号)
End Sub

Private Sub txt医保号_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intinsure As Integer
    Dim intOrder As Integer
    Dim str交易号 As String
    Dim str入参1 As String
    Dim str入参2 As String
    Dim str出参 As String
    On Error GoTo errHand
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    intinsure = Me.cbo子医保.ItemData(Me.cbo子医保.ListIndex)
    If Not CreateObject_Insure(intinsure, intOrder) Then Exit Sub
    If Not gobjInsure_Obj(intOrder).InitInsure(gcnOracle, intinsure) Then Exit Sub
    
    '身份验证
    str交易号 = "01"
    str入参1 = Trim(txt医保号.Text)
    If gobjInsure_Obj(intOrder).CallAPI(str交易号, str入参1, str入参2, str出参) Then
        '0证号|1姓名|2性别|3出生日期|4身份证号|5余额|6类别|7家庭住址|8邮编|9社委会
        txt姓名.Text = Split(str出参, "|")(1)
        txt帐户余额.Text = Val(Split(str出参, "|")(5))
        txt帐户余额.Tag = Val(Split(str出参, "|")(10))
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
