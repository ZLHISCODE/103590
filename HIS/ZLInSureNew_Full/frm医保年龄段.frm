VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm医保年龄段 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "保险人群年龄段"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6180
   Icon            =   "frm医保年龄段.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   6180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CheckBox chk属性 
      Caption         =   "无封顶线(&T)——该类人群无统筹报销封顶线限制"
      Height          =   210
      Index           =   2
      Left            =   825
      TabIndex        =   2
      Top             =   1380
      Width           =   4275
   End
   Begin VB.TextBox txt段数 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   300
      Left            =   1830
      MaxLength       =   1
      TabIndex        =   3
      Text            =   "1"
      Top             =   1845
      Width           =   450
   End
   Begin VB.CheckBox chk属性 
      Caption         =   "无起付线(&S)——该类人群无统筹报销起付线限制"
      Height          =   210
      Index           =   1
      Left            =   825
      TabIndex        =   1
      Top             =   1065
      Width           =   4275
   End
   Begin VB.CheckBox chk属性 
      Caption         =   "全额统筹(&A)——该类人群所有保险项目无首先自付比例"
      Height          =   210
      Index           =   0
      Left            =   825
      TabIndex        =   0
      Top             =   735
      Width           =   4785
   End
   Begin VB.TextBox txtInput 
      Height          =   285
      Left            =   3870
      MaxLength       =   16
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3210
      Visible         =   0   'False
      Width           =   1425
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh分段 
      Height          =   1395
      Left            =   825
      TabIndex        =   5
      Top             =   2160
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   2461
      _Version        =   393216
      Cols            =   4
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
   End
   Begin VB.Frame fraTop 
      Height          =   30
      Left            =   -60
      TabIndex        =   12
      Top             =   585
      Width           =   7125
   End
   Begin VB.Frame fraButtom 
      Height          =   30
      Left            =   -30
      TabIndex        =   11
      Top             =   3690
      Width           =   7125
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   315
      TabIndex        =   10
      Top             =   3825
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3255
      TabIndex        =   8
      Top             =   3825
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4500
      TabIndex        =   9
      Top             =   3825
      Width           =   1100
   End
   Begin MSComCtl2.UpDown ud段数 
      Height          =   300
      Left            =   2280
      TabIndex        =   4
      Top             =   1845
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   529
      _Version        =   393216
      Value           =   1
      AutoBuddy       =   -1  'True
      BuddyControl    =   "txt段数"
      BuddyDispid     =   196610
      OrigLeft        =   2250
      OrigTop         =   1875
      OrigRight       =   2490
      OrigBottom      =   2175
      Max             =   9
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.Label lblSect 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "分段数目(&N)"
      Height          =   180
      Left            =   825
      TabIndex        =   13
      Top             =   1905
      Width           =   990
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "根据相关规则设置分段，以便进一步设置分段支付比例。"
      Height          =   180
      Left            =   840
      TabIndex        =   6
      Top             =   225
      Width           =   4500
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   195
      Picture         =   "frm医保年龄段.frx":000C
      Top             =   60
      Width           =   480
   End
End
Attribute VB_Name = "frm医保年龄段"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mlng险类 As Long, mlng中心 As Long, mlngIndex As Long
Dim mdbl间隔值 As Double
Dim mstrFormat As String   '格式化串
Dim mstr位置 As String
Dim mblnOK As Boolean
Dim mblnChange As Boolean     '是否改变了


Private Sub chk属性_Click(Index As Integer)
    mblnChange = True
End Sub

Private Sub chk属性_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = False Then Exit Sub
    If MsgBox("如果你就这样退出的话，所有的修改都不会生效。" & vbCrLf & "是否确认退出？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
    End If
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name & 1
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
                If Val(.TextMatrix(lngRow, 2)) > 200 Or Val(.TextMatrix(lngRow - 1, 3)) > 200 Then
                    MsgBox "年龄的上下限不能超过200，请检查！", vbInformation, gstrSysName
                    .SetFocus: Exit Sub
                End If
            End If
            .TextMatrix(lngRow, 1) = Join(Split(.TextMatrix(lngRow, 1), ";"))   '去掉可能手工输入的";"
            strSects = strSects & Val(.TextMatrix(lngRow, 0)) & ";" & Trim(.TextMatrix(lngRow, 1)) & ";" & Val(.TextMatrix(lngRow, 2)) & ";" & Val(.TextMatrix(lngRow, 3)) & ";"
        Next
    End With
    
    On Error GoTo errHandle
    
    gstrSQL = "zl_保险年龄段_Update(" & mlng险类 & "," & mlng中心 & "," & mlngIndex & "," & _
            chk属性(0).Value & "," & chk属性(1).Value & "," & chk属性(2).Value & ",'" & strSects & "')"
    
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

Public Function 档次设置(ByVal lng险类 As Long, ByVal lng中心 As Long, ByVal lngIndex As Long, ByVal STRNAME As String) As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim lngRow As Long
        
    mblnOK = False
    mlng险类 = lng险类
    mlng中心 = lng中心
    mlngIndex = lngIndex
    
    Call InitTable
    frm医保年龄段.Caption = STRNAME & "年龄段设置"
    mdbl间隔值 = 1
    gstrSQL = "select 年龄段 as 序号,名称,下限,上限,nvl(全额统筹,0) as 全额统筹 ,nvl(无起付线,0) as 无起付线 ,nvl(无封顶线,0) as 无封顶线 " & _
            " from 保险年龄段" & _
            " where 险类=[1] and 中心=[2] and 在职=[3]" & _
            " Order by 年龄段"
    mstrFormat = "###;-###; ; "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng险类, lng中心, lngIndex)
    
    If rsTemp.EOF Then
        ud段数.Value = 1
        msh分段.Rows = 2
        msh分段.TextMatrix(1, 0) = 1
        msh分段.TextMatrix(1, 1) = STRNAME
    Else
        ud段数.Value = rsTemp.RecordCount
        msh分段.Rows = rsTemp.RecordCount + 1
        
        chk属性(0).Value = IIf(rsTemp("全额统筹") = 1, 1, 0)
        chk属性(1).Value = IIf(rsTemp("无起付线") = 1, 1, 0)
        chk属性(2).Value = IIf(rsTemp("无封顶线") = 1, 1, 0)
        
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
    frm医保年龄段.Show vbModal, frm医保类别
    档次设置 = mblnOK
End Function
