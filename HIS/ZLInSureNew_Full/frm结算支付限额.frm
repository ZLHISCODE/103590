VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm结算支付限额 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "年度支付限额"
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4500
   Icon            =   "frm结算支付限额.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txt住院次数 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   300
      Left            =   1920
      MaxLength       =   2
      TabIndex        =   12
      Text            =   "1"
      Top             =   1020
      Width           =   330
   End
   Begin VB.Frame fraTop 
      Height          =   30
      Left            =   0
      TabIndex        =   10
      Top             =   570
      Width           =   7125
   End
   Begin VB.Frame fraButtom 
      Height          =   30
      Left            =   -195
      TabIndex        =   8
      Top             =   4320
      Width           =   7125
   End
   Begin VB.TextBox txt封顶线 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   1920
      TabIndex        =   0
      Top             =   675
      Width           =   1005
   End
   Begin VB.TextBox txtInput 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3270
      MaxLength       =   16
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3675
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   195
      TabIndex        =   5
      Top             =   4470
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   1995
      TabIndex        =   2
      Top             =   4470
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   3210
      TabIndex        =   3
      Top             =   4470
      Width           =   1100
   End
   Begin MSComCtl2.UpDown ud住院次数 
      Height          =   300
      Left            =   2250
      TabIndex        =   11
      Top             =   1020
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "txt住院次数"
      BuddyDispid     =   196609
      OrigLeft        =   2250
      OrigTop         =   1020
      OrigRight       =   2490
      OrigBottom      =   1320
      Max             =   99
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh起付线 
      Height          =   2775
      Left            =   930
      TabIndex        =   1
      Top             =   1380
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   4895
      _Version        =   393216
      BackColorBkg    =   -2147483643
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Image imgTop 
      Height          =   480
      Left            =   150
      Picture         =   "frm结算支付限额.frx":000C
      Top             =   15
      Width           =   480
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "重庆医疗保险2002年医保基金支付限额规则"
      Height          =   180
      Left            =   780
      TabIndex        =   9
      Top             =   270
      Width           =   3420
   End
   Begin VB.Label lblTop 
      AutoSize        =   -1  'True
      Caption         =   "1)年度封顶线             元"
      Height          =   180
      Left            =   735
      TabIndex        =   7
      Top             =   750
      Width           =   2430
   End
   Begin VB.Label lblStart 
      AutoSize        =   -1  'True
      Caption         =   "2)年度起付线(        次住院起金额不变)"
      Height          =   180
      Left            =   720
      TabIndex        =   4
      Top             =   1080
      Width           =   3420
   End
End
Attribute VB_Name = "frm结算支付限额"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mlng险类 As Long, mlng中心 As Long, mlng年度 As Long
Dim mblnOK As Boolean
Dim mblnChange As Boolean     '是否改变了
Dim mstr位置 As String

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = False Then Exit Sub
    If MsgBox("如果你就这样退出的话，所有的修改都不会生效。" & vbCrLf & "是否确认退出？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim lngRow As Long
    
    Dim strSects As String
    If Val(txt封顶线.Text) <= 0 Then
        MsgBox "封顶线金额必须大于0。", vbInformation, gstrSysName
        txt封顶线.SetFocus
        Exit Sub
    End If
    If Val(txt封顶线.Text) > 10000000 Then
        MsgBox "封顶线金额不能大于1000万。", vbInformation, gstrSysName
        txt封顶线.SetFocus
        Exit Sub
    End If
    
    strSects = "A;" & Val(txt封顶线.Text) & ";"
    With msh起付线
        For lngRow = .FixedRows To .Rows - 1
            If Val(.TextMatrix(lngRow, 1)) <= 0 Or Val(.TextMatrix(lngRow, 1)) > 100000 Then
                MsgBox "第" & lngRow & "次住院起付线金额未设置正确。", vbInformation, gstrSysName
                .SetFocus
                Exit Sub
            End If
            
            If lngRow > .FixedRows And Val(.TextMatrix(lngRow, 1)) > Val(.TextMatrix(lngRow - 1, 1)) Then
                MsgBox "第" & lngRow & "次住院起付线金额比上一次还大。", vbInformation, gstrSysName
                .SetFocus
                Exit Sub
            End If
            strSects = strSects & .TextMatrix(lngRow, 0) & ";" & Val(.TextMatrix(lngRow, 1)) & ";"
        Next
    End With
    
    On Error GoTo ErrHand
    gstrSQL = "zl_保险支付限额_Update(" & mlng险类 & "," & mlng中心 & "," & mlng年度 & ",'" & strSects & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    mblnOK = True
    mblnChange = False
    Unload Me
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub msh起付线_DblClick()
    With msh起付线
        If .Col = 0 Then Exit Sub
        txtInput.Alignment = 1
        txtInput.Move .Left + .CellLeft - 15, .Top + .CellTop - 15, .ColWidth(.Col) - 15, .RowHeight(.Row) - 15
        txtInput.Text = .TextMatrix(.Row, .Col)
        mstr位置 = .Row & ";" & .Col
        txtInput.Visible = True
        txtInput.ZOrder
        zlControl.TxtSelAll txtInput
        txtInput.SetFocus
    End With
End Sub

Private Sub msh起付线_KeyPress(KeyAscii As Integer)
    With msh起付线
        Select Case KeyAscii
        Case 13                 'Enter
            If .Col = .Cols - 1 Then
                If .Row = .Rows - 1 Then
                    '离开网格
                    Me.cmdOK.SetFocus
                    Exit Sub
                End If
                '下一行
                .Row = .Row + 1
                .Col = .FixedCols
                .TopRow = .Row
            Else
                '后一列
                .Col = .Col + 1
            End If
        Case 27                     'ESC退出
            Call cmdCancel_Click
        Case 32                     '空格键进入编辑
            Call msh起付线_DblClick
        Case Else                   '其他直接进入编辑
            Call msh起付线_DblClick
            If .Col = 1 And (KeyAscii = 46 Or KeyAscii >= 48 And KeyAscii <= 64) Then
                '数字键进入编辑
                txtInput.Text = Chr(KeyAscii)
                txtInput.SelStart = Len(txtInput.Text)
            End If
        End Select
    End With
End Sub

Private Sub msh起付线_RowColChange()
    msh起付线.TopRow = msh起付线.Row
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)
    Dim lngRow As Long, lngCol As Long
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call txtInput_Validate(False)
        If txtInput.Visible Then
            Exit Sub
        Else
            msh起付线.SetFocus
            Call msh起付线_KeyPress(vbKeyReturn)
        End If
    ElseIf KeyAscii = vbKeyEscape Then
        KeyAscii = 0
        lngRow = Split(mstr位置, ";")(0)
        lngCol = Split(mstr位置, ";")(1)
        txtInput.Text = msh起付线.TextMatrix(lngRow, lngCol)
        txtInput.Visible = False
        msh起付线.SetFocus
    ElseIf KeyAscii = 45 Or KeyAscii = 47 Or KeyAscii > 65 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtInput_LostFocus()
    txtInput.Visible = False
End Sub

Private Sub txtInput_Validate(Cancel As Boolean)
    Dim lngRow As Long, lngCol As Long
    
    With msh起付线
        If txtInput.Visible = False Then Exit Sub
        lngRow = Split(mstr位置, ";")(0)
        lngCol = Split(mstr位置, ";")(1)
        If Val(txtInput.Text) <= 0 Or Val(txtInput.Text) > 100000 Then
            MsgBox "起付线金额必须大于0且小于10万。", vbInformation, gstrSysName
            DoEvents
            Cancel = True
            txtInput.Visible = True
            txtInput.SetFocus
            Exit Sub
        End If
        '填写单元数值
        mblnChange = True
        .TextMatrix(lngRow, lngCol) = Format(Val(txtInput.Text), "0")
        txtInput.Visible = False
    End With

End Sub

Private Sub txt住院次数_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub txt封顶线_Change()
    mblnChange = True
End Sub

Private Sub txt封顶线_GotFocus()
    zlControl.TxtSelAll txt封顶线
End Sub

Private Sub txt封顶线_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub txt封顶线_LostFocus()
    txt封顶线.Text = Format(txt封顶线.Text, "0")
End Sub

Private Sub txt封顶线_Validate(Cancel As Boolean)
    If Val(txt封顶线.Text) <= 0 Then
        Cancel = True
        MsgBox "封顶线金额必须大于0。", vbInformation, gstrSysName
        Exit Sub
    End If
End Sub

Private Sub ud住院次数_Change()
    Dim lngRow As Long
    
    With msh起付线
        .Rows = ud住院次数.Value + 1
        lngRow = .Rows - 1
        .TextMatrix(lngRow, 0) = lngRow
        If Trim(.TextMatrix(lngRow, 1)) = "" Then
            .TextMatrix(lngRow, 1) = 0
        End If
    End With
End Sub

Public Function 编辑支付限额(ByVal lng险类 As Long, ByVal lng中心 As Long, ByVal lng年度 As Long) As Boolean
'功能:用来与调用的窗口进行通讯的程序
'返回值:编辑成功返回True,否则为False
    
    mlng险类 = lng险类
    mlng中心 = lng中心
    mlng年度 = lng年度
    mblnOK = False
    
    Dim lngRow As Long
    
    lblNote.Caption = lng年度 & "年度医保基金支付限额规则"
    
    With msh起付线
        .TextMatrix(0, 0) = "住院次数"
        .TextMatrix(0, 1) = "金额"
        .ColAlignmentFixed(0) = 4
        .ColAlignmentFixed(1) = 4
        .ColAlignment(0) = 4
        .ColAlignment(1) = 7
        .ColWidth(0) = 1000
        .ColWidth(1) = 1200
    End With
    With frm结算规则.msh支付限额
        txt封顶线.Text = .TextMatrix(.Rows - 1, 1)
        ud住院次数.Value = .Rows - 2
        msh起付线.Rows = .Rows - 1
        
        For lngRow = 1 To .Rows - 2
            msh起付线.TextMatrix(lngRow, 0) = lngRow
            msh起付线.TextMatrix(lngRow, 1) = .TextMatrix(lngRow, 1)
        Next
    End With
    
    mblnChange = False
    frm结算支付限额.Show vbModal, frm结算规则
    编辑支付限额 = mblnOK
End Function
