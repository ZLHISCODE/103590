VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm报销支付比例 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "年度支付比例"
   ClientHeight    =   5445
   ClientLeft      =   1770
   ClientTop       =   2235
   ClientWidth     =   8250
   Icon            =   "frm报销支付比例.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   8250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraButtom 
      Height          =   30
      Left            =   -195
      TabIndex        =   5
      Top             =   4845
      Width           =   9345
   End
   Begin VB.TextBox txtInput 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3660
      MaxLength       =   16
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2520
      Visible         =   0   'False
      Width           =   1425
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh比例 
      Height          =   4035
      Left            =   330
      TabIndex        =   0
      Top             =   690
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   7117
      _Version        =   393216
      BackColorBkg    =   -2147483643
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   195
      TabIndex        =   3
      Top             =   4995
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   5730
      TabIndex        =   1
      Top             =   4995
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6945
      TabIndex        =   2
      Top             =   4995
      Width           =   1100
   End
   Begin VB.Image imgTop 
      Height          =   480
      Left            =   150
      Picture         =   "frm报销支付比例.frx":000C
      Top             =   15
      Width           =   480
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "重庆医疗保险2002年医保基金支付分段比例规则"
      Height          =   180
      Left            =   780
      TabIndex        =   6
      Top             =   300
      Width           =   3780
   End
End
Attribute VB_Name = "frm报销支付比例"
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
    Dim lng年龄段 As Long, lngPreRow As Long
    
    Dim strSects As String
    
    strSects = ""
    With msh比例
        For lngRow = .FixedRows To .Rows - 1
            For lngCol = .FixedCols To .Cols - 1
                If Val(.TextMatrix(lngRow, lngCol)) < 0 Or Val(.TextMatrix(lngRow, lngCol)) > 100 Then
                    MsgBox .TextMatrix(lngRow, 0) & .TextMatrix(0, lngCol) & "的比例未设置正确。", vbInformation, gstrSysName
                    .Row = lngRow: .Col = lngCol: .SetFocus: Exit Sub
                End If
            Next
        Next
    End With
    
    On Error GoTo ErrHand
    gcnOracle.BeginTrans
    gstrSQL = "ZL_保险报销政策_DELETE(1," & mlng险类 & "," & mlng中心 & "," & mlng年度 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    With msh比例
        For lngRow = .FixedRows To .Rows - 1
            For lngCol = .FixedCols To .Cols - 1
                gstrSQL = "ZL_保险报销政策_INSERT(1," & mlng险类 & "," & mlng中心 & "," & mlng年度 & _
                "," & IIf(.TextMatrix(1, lngCol) = "本院", 1, 2) & "," & .ColData(lngCol) & _
                "," & .RowData(lngRow) & "," & Val(.TextMatrix(lngRow, lngCol)) & ")"
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

Private Sub msh比例_DblClick()
    With msh比例
        If .Col = 0 Then Exit Sub
        txtInput.Alignment = 1
        txtInput.Move .Left + .CellLeft - 15, .Top + .CellTop - 15, .ColWidth(.Col) - 15, .RowHeight(.Row) - 15
        txtInput.Text = .TextMatrix(.Row, .Col)
        mstr位置 = .Row & ";" & .Col
        txtInput.Visible = True
        zlControl.TxtSelAll txtInput
        txtInput.SetFocus
    End With
End Sub

Private Sub msh比例_KeyPress(KeyAscii As Integer)
    With msh比例
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
            Call msh比例_DblClick
        Case Else                   '其他直接进入编辑
            Call msh比例_DblClick
            If .Col <> 0 And (KeyAscii = 46 Or KeyAscii >= 48 And KeyAscii <= 64) Then
                '数字键进入编辑
                txtInput.Text = Chr(KeyAscii)
                txtInput.SelStart = Len(txtInput.Text)
            End If
        End Select
    End With
End Sub

Private Sub msh比例_RowColChange()
    msh比例.TopRow = msh比例.Row
    msh比例.LeftCol = msh比例.Col
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)
    Dim lngRow As Long, lngCol As Long

    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call txtInput_Validate(False)
        If txtInput.Visible Then
            Exit Sub
        Else
            msh比例.SetFocus
            Call msh比例_KeyPress(13)
        End If
    ElseIf KeyAscii = vbKeyEscape Then
        KeyAscii = 0
        lngRow = Split(mstr位置, ";")(0)
        lngCol = Split(mstr位置, ";")(1)
        txtInput.Text = msh比例.TextMatrix(lngRow, lngCol)
        txtInput.Visible = False
        msh比例.SetFocus
    ElseIf KeyAscii = 45 Or KeyAscii = 47 Or KeyAscii > 65 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtInput_LostFocus()
    txtInput.Visible = False
End Sub

Private Sub txtInput_Validate(Cancel As Boolean)
    Dim lngRow As Long, lngCol As Long

    With msh比例
        If txtInput.Visible = False Then Exit Sub
        lngRow = Split(mstr位置, ";")(0)
        lngCol = Split(mstr位置, ";")(1)
        If Val(txtInput.Text) < 0 Or Val(txtInput.Text) > 100 Then
            MsgBox "比例必须大于等于0且小于等于100。", vbInformation, gstrSysName
            DoEvents
            Cancel = True
            txtInput.Visible = True
            zlControl.TxtSelAll txtInput
            txtInput.SetFocus
            Exit Sub
        End If
        '填写单元数值
        mblnChange = True
        .TextMatrix(lngRow, lngCol) = Format(Val(txtInput.Text), "0.00")
        txtInput.Visible = False
    End With
End Sub

Public Function 编辑支付比例(ByVal lng险类 As Long, ByVal lng中心 As Long, ByVal lng年度 As Long) As Boolean
'功能:用来与调用的窗口进行通讯的程序
'返回值:编辑成功返回True,否则为False
    
    mlng险类 = lng险类
    mlng中心 = lng中心
    mlng年度 = lng年度
    mblnOK = False
    
    Dim lngCount As Integer, lngRow As Long, lngCol As Long
    
    lblNote.Caption = lng年度 & "年度医保基金支付分段比例（百分比）规则"
    With frm保险报销政策.msh支付比例
        msh比例.Rows = .Rows
        msh比例.Cols = .Cols
        msh比例.FixedRows = .FixedRows
        msh比例.ColWidth(0) = .ColWidth(0)
        msh比例.ColAlignmentFixed(0) = 1
        msh比例.Row = 0
        msh比例.Col = 0
        msh比例.CellAlignment = 4
        For lngCol = msh比例.FixedCols To .Cols - 1
            msh比例.ColData(lngCol) = .ColData(lngCol)
            msh比例.ColWidth(lngCol) = .ColWidth(lngCol)
            msh比例.ColAlignmentFixed(lngCol) = 4
            msh比例.ColAlignment(lngCol) = 7
        Next
        For lngRow = 0 To .Rows - 1
            msh比例.RowData(lngRow) = .RowData(lngRow)
            For lngCol = 0 To .Cols - 1
                msh比例.TextMatrix(lngRow, lngCol) = .TextMatrix(lngRow, lngCol)
            Next
        Next
        msh比例.Row = .FixedRows
        msh比例.Col = .FixedCols
        
        msh比例.MergeCells = flexMergeFree
        msh比例.MergeRow(0) = True
        msh比例.MergeCol(0) = True
    End With
    
    mblnChange = False
    frm报销支付比例.Show vbModal, frm保险报销政策
    编辑支付比例 = mblnOK
End Function

