VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmMoneyEnum 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "现金点钞"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4620
   Icon            =   "frmMoneyEnum.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Height          =   180
      IMEMode         =   3  'DISABLE
      Left            =   3060
      MaxLength       =   10
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1125
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txt大写 
      Height          =   300
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5460
      Width           =   3885
   End
   Begin VB.TextBox txt合计 
      Height          =   300
      Left            =   600
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5025
      Width           =   3885
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshMoney 
      Height          =   4080
      Left            =   120
      TabIndex        =   0
      Top             =   765
      Width           =   4365
      _ExtentX        =   7699
      _ExtentY        =   7197
      _Version        =   393216
      BackColor       =   16777215
      Rows            =   14
      Cols            =   3
      FixedCols       =   0
      RowHeightMin    =   280
      BackColorBkg    =   16777215
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      GridLinesFixed  =   1
      ScrollBars      =   2
      AllowUserResizing=   1
      FormatString    =   "^      人民币      |^  面额(元)  |^    张数    "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmMoneyEnum.frx":058A
      _NumberOfBands  =   1
      _Band(0).Cols   =   3
   End
   Begin VB.Label Label3 
      Caption         =   "请根据持有现钞情况，在不同的人民币面额中输入钞票张数，系统自动计算现金合计。"
      Height          =   360
      Left            =   750
      TabIndex        =   5
      Top             =   210
      Width           =   3420
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "大写"
      Height          =   180
      Left            =   165
      TabIndex        =   3
      Top             =   5520
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "合计"
      Height          =   180
      Left            =   165
      TabIndex        =   1
      Top             =   5085
      Width           =   360
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   135
      Picture         =   "frmMoneyEnum.frx":08A4
      Top             =   165
      Width           =   480
   End
End
Attribute VB_Name = "frmMoneyEnum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mdblMoney As Double
Public Sub ShowMe(frmMain As Object, Optional dblMoney As Double)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:程序入口
    '入参:frmMain -调用的主窗体
    '        dblMoney-当前要清点的金额
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-09-13 16:06:28
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mdblMoney = dblMoney
    On Error Resume Next
    Me.Show 1, frmMain
End Sub

Private Sub Form_Activate()
    If Not txtEdit.Visible Then Call mshMoney_EnterCell
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then Unload Me
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    
    On Error GoTo errH
    strSQL = "Select 名称,面额 From 人民币面额 Order by 编码"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    If rsTmp.EOF Then
        MsgBox "请先到字典管理中设置人民币面额。", vbInformation, gstrSysName
        Unload Me: Exit Sub
    End If
        
    mshMoney.Rows = rsTmp.RecordCount + 1
    For i = 1 To rsTmp.RecordCount
        mshMoney.TextMatrix(i, 0) = rsTmp!名称
        mshMoney.TextMatrix(i, 1) = Format(Nvl(rsTmp!面额, 0), "0.00")
        rsTmp.MoveNext
    Next
    
    mshMoney.ColAlignment(0) = 1
    mshMoney.ColAlignment(1) = 1
    mshMoney.ColAlignment(2) = 1
    mshMoney.Row = 1: mshMoney.Col = 2
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mshMoney_EnterCell()
    If mshMoney.Col = 2 Then
        txtEdit.Left = mshMoney.Left + mshMoney.CellLeft + 15
        txtEdit.Top = mshMoney.Top + mshMoney.CellTop + (mshMoney.CellHeight - txtEdit.Height) / 2
        txtEdit.Width = mshMoney.CellWidth - 30
        txtEdit.Text = mshMoney.Text
        txtEdit.Visible = True
        txtEdit.ZOrder
        If Visible Then txtEdit.SetFocus
        mshMoney.CellBackColor = txtEdit.BackColor
    Else
        txtEdit.Visible = False
        mshMoney.CellBackColor = mshMoney.BackColor
    End If
End Sub

Private Sub mshMoney_LeaveCell()
    txtEdit.Visible = False
    mshMoney.CellBackColor = mshMoney.BackColor
End Sub

Private Sub mshMoney_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If mshMoney.MouseRow >= mshMoney.FixedRows Then
        Call mshMoney_EnterCell
    End If
End Sub

Private Sub mshMoney_Scroll()
    txtEdit.Visible = False
    mshMoney.CellBackColor = mshMoney.BackColor
End Sub

Private Sub txtEdit_GotFocus()
    Call zlControl.TxtSelAll(txtEdit)
End Sub

Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 13, vbKeyDown
            If mshMoney.Row + 1 <= mshMoney.Rows - 1 Then
                Call txtEdit_Validate(False)
                Call mshMoney_LeaveCell
                mshMoney.Row = mshMoney.Row + 1
                Call mshMoney_EnterCell
            End If
        Case vbKeyUp
            If mshMoney.Row - 1 >= mshMoney.FixedRows Then
                Call txtEdit_Validate(False)
                Call mshMoney_LeaveCell
                mshMoney.Row = mshMoney.Row - 1
                Call mshMoney_EnterCell
            End If
    End Select
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtEdit_Validate(Cancel As Boolean)
    mshMoney.Text = IIf(Val(txtEdit.Text) = 0, "", Val(txtEdit.Text))
    txtEdit.Visible = False
    mshMoney.CellBackColor = mshMoney.BackColor
    
    Call CalcMoney
End Sub

Private Sub txt大写_GotFocus()
    Call zlControl.TxtSelAll(txt大写)
End Sub

Private Sub txt合计_GotFocus()
    Call zlControl.TxtSelAll(txt合计)
End Sub

Private Sub CalcMoney()
    Dim curMoney As Currency, i As Long
    
    For i = 1 To mshMoney.Rows - 1
        curMoney = curMoney + Val(mshMoney.TextMatrix(i, 1)) * Val(mshMoney.TextMatrix(i, 2))
    Next
    If curMoney <> 0 Then
        txt合计.Text = Format(curMoney, "0.00") & "(元)"
        txt大写.Text = zlCommFun.UppeMoney(curMoney)
    Else
        txt合计.Text = "": txt大写.Text = ""
    End If
End Sub
