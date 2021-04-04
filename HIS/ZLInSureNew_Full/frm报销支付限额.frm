VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm报销支付限额 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "年度支付限额"
   ClientHeight    =   5430
   ClientLeft      =   2925
   ClientTop       =   3660
   ClientWidth     =   8235
   Icon            =   "frm报销支付限额.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   8235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txt住院次数 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   300
      Left            =   1560
      MaxLength       =   2
      TabIndex        =   10
      Text            =   "1"
      Top             =   720
      Width           =   330
   End
   Begin VB.Frame fraTop 
      Height          =   30
      Left            =   0
      TabIndex        =   8
      Top             =   570
      Width           =   7125
   End
   Begin VB.Frame fraButtom 
      Height          =   30
      Left            =   -195
      TabIndex        =   6
      Top             =   4830
      Width           =   8565
   End
   Begin VB.TextBox txtInput 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2910
      MaxLength       =   16
      TabIndex        =   5
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
      TabIndex        =   4
      Top             =   4980
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   5715
      TabIndex        =   1
      Top             =   4980
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6930
      TabIndex        =   2
      Top             =   4980
      Width           =   1100
   End
   Begin MSComCtl2.UpDown ud住院次数 
      Height          =   300
      Left            =   1890
      TabIndex        =   9
      Top             =   720
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
      Height          =   3525
      Left            =   360
      TabIndex        =   0
      Top             =   1140
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   6218
      _Version        =   393216
      BackColorBkg    =   -2147483643
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Image imgTop 
      Height          =   480
      Left            =   150
      Picture         =   "frm报销支付限额.frx":000C
      Top             =   15
      Width           =   480
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "重庆医疗保险2002年医保基金支付限额规则"
      Height          =   180
      Left            =   780
      TabIndex        =   7
      Top             =   270
      Width           =   3420
   End
   Begin VB.Label lblStart 
      AutoSize        =   -1  'True
      Caption         =   "1)年度起付线(        次住院起金额不变)"
      Height          =   180
      Left            =   360
      TabIndex        =   3
      Top             =   780
      Width           =   3420
   End
End
Attribute VB_Name = "frm报销支付限额"
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
    Dim lngRow As Long, lngCol As Long
    
    With msh起付线
        For lngRow = .FixedRows To .Rows - 1
            If lngRow <> .Rows - 1 Then
                If Val(.TextMatrix(lngRow, 1)) < 0 Or Val(.TextMatrix(lngRow, 1)) > 100000 Then
                    MsgBox "第" & lngRow & "次住院起付线金额未设置正确。", vbInformation, gstrSysName
                    .SetFocus
                    Exit Sub
                End If
                If lngRow > .FixedRows And Val(.TextMatrix(lngRow, 1)) > Val(.TextMatrix(lngRow - 1, 1)) Then
                    MsgBox "第" & lngRow & "次住院起付线金额比上一次还大。", vbInformation, gstrSysName
                    .SetFocus
                    Exit Sub
                End If
            Else
                If Val(.TextMatrix(lngRow, 1)) < 0 Or Val(.TextMatrix(lngRow, 1)) > 100000 Then
                    MsgBox "封顶线金额未设置正确。", vbInformation, gstrSysName
                    .SetFocus
                    Exit Sub
                End If
            End If
        Next
    End With
    
    On Error GoTo ErrHand
    gcnOracle.BeginTrans
    gstrSQL = "ZL_保险报销政策_DELETE(2," & mlng险类 & "," & mlng中心 & "," & mlng年度 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    With msh起付线
        For lngRow = .FixedRows To .Rows - 1
            For lngCol = .FixedCols To .Cols - 1
                If lngRow = .Rows - 1 Then
                    '封顶线
                    gstrSQL = "ZL_保险报销政策_INSERT(2," & mlng险类 & "," & mlng中心 & "," & mlng年度 & _
                    "," & IIf(.TextMatrix(1, lngCol) = "本院", 1, 2) & "," & .ColData(lngCol) & _
                    ",1,0,'A'," & Val(.TextMatrix(lngRow, lngCol)) & ")"
                Else
                    gstrSQL = "ZL_保险报销政策_INSERT(2," & mlng险类 & "," & mlng中心 & "," & mlng年度 & _
                    "," & IIf(.TextMatrix(1, lngCol) = "本院", 1, 2) & "," & .ColData(lngCol) & _
                    ",1,0," & lngRow - 1 & "," & Val(.TextMatrix(lngRow, lngCol)) & ")"
                End If
                Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            Next
        Next
    End With
    gcnOracle.CommitTrans
    
    mblnOK = True
    mblnChange = False
    Unload Me
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    gcnOracle.RollbackTrans
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
            If (KeyAscii = 46 Or KeyAscii >= 48 And KeyAscii <= 64) Then
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
        If Val(txtInput.Text) < 0 Or Val(txtInput.Text) > 100000 Then
            MsgBox "起付线金额必须大于等于0且小于10万。", vbInformation, gstrSysName
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

Private Sub ud住院次数_Change()
    Dim lngRow As Long, lngCol As Long
    Dim blnAdd As Boolean
    
    With msh起付线
        If .Rows < ud住院次数.Value + 3 Then
            blnAdd = True
        ElseIf .Rows > ud住院次数.Value + 3 Then
            blnAdd = False
            For lngCol = 1 To .Cols - 1
                .TextMatrix(.Rows - 2, lngCol) = .TextMatrix(.Rows - 1, lngCol)
            Next
        Else
            Exit Sub
        End If
        .Rows = ud住院次数.Value + 3
        lngRow = .Rows - 2
        .TextMatrix(lngRow, 0) = lngRow - 1
        .TextMatrix(.Rows - 1, 0) = "封顶线"
        
        '最后一行下调
        If blnAdd Then
            For lngCol = 1 To .Cols - 1
                .TextMatrix(.Rows - 1, lngCol) = .TextMatrix(.Rows - 2, lngCol)
                .TextMatrix(.Rows - 2, lngCol) = ""
            Next
        End If
    End With
End Sub

Public Function 编辑支付限额(ByVal lng险类 As Long, ByVal lng中心 As Long, ByVal lng年度 As Long) As Boolean
'功能:用来与调用的窗口进行通讯的程序
'返回值:编辑成功返回True,否则为False
    Dim lngRow As Long, lngCol As Long
    mlng险类 = lng险类
    mlng中心 = lng中心
    mlng年度 = lng年度
    mblnOK = False
    
    lblNote.Caption = lng年度 & "年度医保基金支付限额规则"
    
    With msh起付线
        .TextMatrix(0, 0) = "住院次数"
        .TextMatrix(0, 1) = "金额"
        .TextMatrix(1, 0) = "1"
        .ColAlignmentFixed(0) = 4
        .ColAlignmentFixed(1) = 4
        .ColAlignment(0) = 4
        .ColAlignment(1) = 7
        .ColWidth(0) = 1000
        .ColWidth(1) = 1200
    End With
    
    With frm保险报销政策.msh支付限额
        ud住院次数.Value = .Rows - 3
        msh起付线.Rows = .Rows
        msh起付线.Cols = .Cols
        msh起付线.FixedRows = .FixedRows
        msh起付线.ColWidth(0) = .ColWidth(0)
        msh起付线.ColAlignmentFixed(0) = 1
        msh起付线.Row = 0
        msh起付线.Col = 0
        msh起付线.CellAlignment = 4
        For lngCol = 0 To .Cols - 1
            msh起付线.ColData(lngCol) = .ColData(lngCol)
            msh起付线.ColWidth(lngCol) = .ColWidth(lngCol)
            msh起付线.ColAlignmentFixed(lngCol) = 4
            msh起付线.ColAlignment(lngCol) = 7
        Next
        For lngRow = 0 To .Rows - 1
            msh起付线.RowData(lngRow) = .RowData(lngRow)
            If lngRow <> .Rows - 1 And lngRow > 1 Then
                msh起付线.TextMatrix(lngRow, 0) = lngRow - 1
            Else
                msh起付线.TextMatrix(lngRow, 0) = .TextMatrix(lngRow, 0)
            End If
            For lngCol = 1 To .Cols - 1
                msh起付线.TextMatrix(lngRow, lngCol) = .TextMatrix(lngRow, lngCol)
            Next
        Next
        msh起付线.Row = .FixedRows
        msh起付线.Col = .FixedCols
        
        msh起付线.MergeCells = flexMergeFree
        msh起付线.MergeRow(0) = True
        msh起付线.MergeCol(0) = True
    End With
    
    mblnChange = False
    frm报销支付限额.Show vbModal, frm保险报销政策
    编辑支付限额 = mblnOK
End Function


