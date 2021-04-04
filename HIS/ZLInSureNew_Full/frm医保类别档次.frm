VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm医保类别档次 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "分段设置"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5445
   Icon            =   "frm医保类别档次.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   5445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox txtInput 
      Height          =   285
      Left            =   3480
      MaxLength       =   16
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2040
      Visible         =   0   'False
      Width           =   1425
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh分段 
      Height          =   1740
      Left            =   570
      TabIndex        =   4
      Top             =   1035
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   3069
      _Version        =   393216
      Cols            =   4
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
   End
   Begin VB.Frame fraTop 
      Height          =   30
      Left            =   -60
      TabIndex        =   10
      Top             =   585
      Width           =   7125
   End
   Begin VB.Frame fraButtom 
      Height          =   30
      Left            =   -60
      TabIndex        =   9
      Top             =   2835
      Width           =   7125
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   285
      TabIndex        =   8
      Top             =   2970
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   2865
      TabIndex        =   6
      Top             =   2970
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4110
      TabIndex        =   7
      Top             =   2970
      Width           =   1100
   End
   Begin MSComCtl2.UpDown ud段数 
      Height          =   300
      Left            =   2070
      TabIndex        =   3
      Top             =   690
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "txt段数"
      BuddyDispid     =   196617
      OrigLeft        =   2085
      OrigTop         =   690
      OrigRight       =   2325
      OrigBottom      =   990
      Max             =   9
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txt段数 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   300
      Left            =   1680
      MaxLength       =   1
      TabIndex        =   2
      Text            =   "1"
      Top             =   690
      Width           =   390
   End
   Begin VB.Label lblSect 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "分段数目(&N)"
      Height          =   180
      Left            =   570
      TabIndex        =   1
      Top             =   750
      Width           =   990
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "根据相关规则设置分段，以便进一步设置分段支付比例。"
      Height          =   180
      Left            =   840
      TabIndex        =   0
      Top             =   225
      Width           =   4500
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   195
      Picture         =   "frm医保类别档次.frx":000C
      Top             =   60
      Width           =   480
   End
End
Attribute VB_Name = "frm医保类别档次"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mlng险类 As Long, mlng中心 As Long
Dim mdbl间隔值 As Double
Dim mstrFormat As String   '格式化串
Dim mstr位置 As String
Dim mblnOK As Boolean
Dim mblnChange As Boolean     '是否改变了


Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = False Then Exit Sub
    If MsgBox("如果你就这样退出的话，所有的修改都不会生效。" & vbCrLf & "是否确认退出？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
    End If
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name & 3
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim strSects As String
    Dim lngRow As Long
    
    With msh分段
        For lngRow = .FixedRows To .Rows - 1
            If Trim(.TextMatrix(lngRow, 1)) = "" Then
                MsgBox "第" & lngRow & "段名称未设置。", vbInformation, gstrSysName
                .SetFocus: Exit Sub
            End If
            If lngRow < .Rows - 1 Then
                If Val(.TextMatrix(lngRow, 2)) > Val(.TextMatrix(lngRow, 3)) Then
                    MsgBox "第" & lngRow & "段上限值应大于下限。", vbInformation, gstrSysName
                    .SetFocus: Exit Sub
                End If
            End If
            If lngRow > .FixedRows Then
                If Val(.TextMatrix(lngRow, 2)) <> Val(.TextMatrix(lngRow - 1, 3)) + mdbl间隔值 Then
                    MsgBox "第" & lngRow & "段下限与上一段上限不连续。", vbInformation, gstrSysName
                    .SetFocus: Exit Sub
                End If
            End If
            If Val(.TextMatrix(lngRow, 2)) <> 0 And Val(.TextMatrix(lngRow - 1, 3)) <> 0 Then
                If Val(.TextMatrix(lngRow, 2)) > 1000000 Or Val(.TextMatrix(lngRow - 1, 3)) > 1000000 Then
                    MsgBox "费用档的上下限不能超过100万，请检查！", vbInformation, gstrSysName
                    .SetFocus: Exit Sub
                End If
            End If
            
            .TextMatrix(lngRow, 1) = Join(Split(.TextMatrix(lngRow, 1), ";"))   '去掉可能手工输入的";"
            strSects = strSects & Val(.TextMatrix(lngRow, 0)) & ";" & Trim(.TextMatrix(lngRow, 1)) & ";" & Val(.TextMatrix(lngRow, 2)) & ";" & Val(.TextMatrix(lngRow, 3)) & ";"
        Next
    End With
    
    On Error GoTo errHandle
    If mlng险类 = TYPE_四川眉山 Then strSects = "0;门诊;0;0;" & strSects
    gstrSQL = "zl_保险费用档_Update(" & mlng险类 & "," & mlng中心 & ",'" & strSects & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    mblnChange = False
    mblnOK = True
    Unload Me
    Exit Sub

errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub InitTable()
    With msh分段
        .TextMatrix(0, 0) = "段"
        .TextMatrix(0, 1) = "名称"
        .TextMatrix(0, 2) = "下限"
        .TextMatrix(0, 3) = "上限"
        .ColAlignmentFixed(0) = 4
        .ColAlignmentFixed(1) = 4
        .ColAlignmentFixed(2) = 4
        .ColAlignmentFixed(3) = 4
        .ColAlignment(0) = 4
        .ColAlignment(1) = 1
        .ColAlignment(2) = 7
        .ColAlignment(3) = 7
        .ColWidth(0) = 300
        .ColWidth(1) = 1800
        .ColWidth(2) = 1000
        .ColWidth(3) = 1000
    End With
End Sub

Private Sub msh分段_DblClick()
    With msh分段
        If .COL = 1 Then txtInput.Alignment = 0
        If .COL = 3 Then txtInput.Alignment = 1
        If .COL = 1 Or .COL = 3 And .Row <> .Rows - 1 Then
            txtInput.Move .Left + .CellLeft - 15, .Top + .CellTop - 15, .ColWidth(.COL) - 15, .RowHeight(.Row) - 15
            txtInput.Text = .TextMatrix(.Row, .COL)
            mstr位置 = .Row & ";" & .COL
            txtInput.Visible = True
            zlControl.TxtSelAll txtInput
            txtInput.SetFocus
        End If
    End With
End Sub

Private Sub msh分段_KeyPress(KeyAscii As Integer)
    With msh分段
        Select Case KeyAscii
        Case 13                 'Enter
            If .COL = .Cols - 1 Then
                If .Row = .Rows - 1 Then
                    '离开网格
                    Me.cmdOK.SetFocus
                    Exit Sub
                End If
                '下一行
                .Row = .Row + 1
                .COL = .FixedCols
                .TopRow = .Row
            Else
                '后一列
                .COL = .COL + 1
            End If
        Case 27                     'ESC退出
            Call cmdCancel_Click
        Case 32                     '空格键进入编辑
            Call msh分段_DblClick
        Case Else                   '其他直接进入编辑
            Call msh分段_DblClick
            If .COL = 1 Then
                txtInput.Text = Chr(KeyAscii)
                txtInput.SelStart = Len(txtInput.Text)
            ElseIf .COL = 3 And (KeyAscii = 46 Or KeyAscii >= 48 And KeyAscii <= 64) Then
                '数字键进入编辑
                txtInput.Text = Chr(KeyAscii)
                txtInput.SelStart = Len(txtInput.Text)
            End If
        End Select
    End With
End Sub

Private Sub msh分段_RowColChange()
    msh分段.TopRow = msh分段.Row
End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)
    Dim lngRow As Long, lngCol As Long
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call txtInput_Validate(False)
        If txtInput.Visible Then
            Exit Sub
        Else
            msh分段.SetFocus
            Call msh分段_KeyPress(13)
        End If
    ElseIf KeyAscii = 27 Then
        KeyAscii = 0
        lngRow = Split(mstr位置, ";")(0)
        lngCol = Split(mstr位置, ";")(1)
        txtInput.Text = msh分段.TextMatrix(lngRow, lngCol)
        txtInput.Visible = False
        msh分段.SetFocus
    Else
        lngCol = Split(mstr位置, ";")(1)
        If lngCol = 3 And (KeyAscii = 45 Or KeyAscii = 47 Or KeyAscii > 65) Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtInput_LostFocus()
    txtInput.Visible = False
End Sub

Private Sub txtInput_Validate(Cancel As Boolean)
    Dim lngRow As Long, lngCol As Long
    
    With msh分段
        If txtInput.Visible = False Then Exit Sub
        lngRow = Split(mstr位置, ";")(0)
        lngCol = Split(mstr位置, ";")(1)
        If lngCol = 3 And Val(txtInput.Text) = 0 Then
            MsgBox "上限不能为0。", vbInformation, gstrSysName
            DoEvents
            Cancel = True
            txtInput.Visible = True
            txtInput.SetFocus
            Exit Sub
        End If
        '填写单元数值
        mblnChange = True
        Select Case lngCol
            Case 1
                .TextMatrix(lngRow, lngCol) = txtInput.Text
            Case 3
                .TextMatrix(lngRow, lngCol) = Format(Val(txtInput.Text), mstrFormat)
                .TextMatrix(lngRow + 1, 2) = Format(Val(.TextMatrix(lngRow, lngCol)) + mdbl间隔值, mstrFormat)
        End Select
        txtInput.Visible = False
    End With

End Sub

Private Sub txt段数_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub ud段数_Change()
    Dim lngRow As Long, lngCol As Long
    
    mblnChange = True
    With msh分段
        .Rows = ud段数.Value + 1
        For lngRow = .FixedRows + 1 To .Rows - 1
            .TextMatrix(lngRow, 0) = lngRow
            If Trim(.TextMatrix(lngRow - 1, 3)) <> "" Then
                .TextMatrix(lngRow, 2) = Format(Val(.TextMatrix(lngRow - 1, 3)) + mdbl间隔值, mstrFormat)
            End If
        Next
        .TextMatrix(.Rows - 1, 3) = ""
    End With
End Sub

Public Function 档次设置(ByVal lng险类 As Long, ByVal lng中心 As Long) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim lngRow As Long
        
    mblnOK = False
    mlng险类 = lng险类
    mlng中心 = lng中心
    Call InitTable
    
    frm医保类别档次.Caption = "支付费用档设置"
    mdbl间隔值 = 0
    mstrFormat = "########0.00;-########0.00; ; "
    
    gstrSQL = "select 档次 as 序号,名称,下限,上限 from 保险费用档 where 险类=[1] and 中心=[2]"
    If mlng险类 = TYPE_四川眉山 Then gstrSQL = gstrSQL & " And 档次<>0"
    gstrSQL = gstrSQL & " Order by 档次"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng险类, lng中心)
    
    If rsTemp.EOF Then
        ud段数.Value = 1
        msh分段.Rows = 2
        msh分段.TextMatrix(1, 0) = 1
        msh分段.TextMatrix(1, 1) = "第一档次"
    Else
        ud段数.Value = rsTemp.RecordCount
        msh分段.Rows = rsTemp.RecordCount + 1
        lngRow = 1
        Do Until rsTemp.EOF
            msh分段.TextMatrix(lngRow, 0) = lngRow
            msh分段.TextMatrix(lngRow, 1) = rsTemp("名称")
            msh分段.TextMatrix(lngRow, 2) = Format(rsTemp("下限"), mstrFormat)
            msh分段.TextMatrix(lngRow, 3) = Format(rsTemp("上限"), mstrFormat)
            
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End If
    
    mblnChange = False
    frm医保类别档次.Show vbModal, frm医保类别
    档次设置 = mblnOK
End Function
