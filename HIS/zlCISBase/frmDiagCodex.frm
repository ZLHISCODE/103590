VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmDiagCodex 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "疾病评估规则"
   ClientHeight    =   7740
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7560
   Icon            =   "frmDiagCodex.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7740
   ScaleWidth      =   7560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   6270
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   900
      Width           =   1200
   End
   Begin MSComctlLib.ListView lvwList 
      Height          =   2925
      Left            =   -5565
      TabIndex        =   27
      Top             =   3885
      Visible         =   0   'False
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   5159
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FlatScrollBar   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imgList"
      SmallIcons      =   "imgList"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "下移(&D)"
      Height          =   350
      Index           =   1
      Left            =   5025
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   5475
      Width           =   1100
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "上移(&U)"
      Height          =   350
      Index           =   0
      Left            =   3930
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5475
      Width           =   1100
   End
   Begin VB.Frame fraDefine 
      Caption         =   "评估细则定义:"
      Height          =   1815
      Left            =   90
      TabIndex        =   25
      Top             =   5895
      Width           =   7395
      Begin VB.ComboBox cboValue 
         Height          =   300
         Left            =   3255
         TabIndex        =   18
         Top             =   915
         Width           =   4035
      End
      Begin VB.ComboBox cboGroup 
         Height          =   300
         Left            =   975
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   285
         Width           =   1920
      End
      Begin VB.TextBox txtGroup 
         Height          =   300
         Left            =   975
         MaxLength       =   10
         TabIndex        =   12
         Top             =   285
         Width           =   1920
      End
      Begin VB.TextBox txtItem 
         Height          =   300
         Left            =   90
         TabIndex        =   14
         Top             =   915
         Width           =   1950
      End
      Begin VB.ComboBox cboFormula 
         Height          =   300
         Left            =   2085
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   915
         Width           =   1140
      End
      Begin VB.TextBox txtDegree 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   990
         MaxLength       =   3
         TabIndex        =   20
         Top             =   1350
         Width           =   1050
      End
      Begin VB.CommandButton cmdAppend 
         Caption         =   "添至细则(&A)"
         Height          =   350
         Left            =   6075
         TabIndex        =   21
         Top             =   1350
         Width           =   1200
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "项目(&I):"
         Height          =   180
         Left            =   90
         TabIndex        =   13
         Top             =   690
         Width           =   720
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         Caption         =   "(通常按疾病分型分期进行分组)"
         Height          =   180
         Left            =   3015
         TabIndex        =   26
         Top             =   345
         Width           =   2520
      End
      Begin VB.Label lblGroup 
         AutoSize        =   -1  'True
         Caption         =   "分组名(&N)"
         Height          =   180
         Left            =   105
         TabIndex        =   10
         Top             =   345
         Width           =   810
      End
      Begin VB.Label lblFormula 
         AutoSize        =   -1  'True
         Caption         =   "条件(&F):"
         Height          =   180
         Left            =   2085
         TabIndex        =   15
         Top             =   690
         Width           =   720
      End
      Begin VB.Label lblValue 
         AutoSize        =   -1  'True
         Caption         =   "值(&V):"
         Height          =   180
         Left            =   3255
         TabIndex        =   17
         Top             =   690
         Width           =   540
      End
      Begin VB.Label lblDegree 
         AutoSize        =   -1  'True
         Caption         =   "怀疑度(&T)"
         Height          =   180
         Left            =   90
         TabIndex        =   19
         Top             =   1410
         Width           =   810
      End
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "移除细则(&R)"
      Height          =   350
      Left            =   6240
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   5475
      Width           =   1200
   End
   Begin VB.CheckBox chkGroup 
      Caption         =   "分组评估(&G)"
      Height          =   240
      Left            =   4110
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1155
      Width           =   1710
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6270
      TabIndex        =   24
      Top             =   480
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6270
      TabIndex        =   23
      Top             =   75
      Width           =   1200
   End
   Begin VB.TextBox txtEnsure 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   2805
      MaxLength       =   5
      TabIndex        =   3
      Text            =   "95"
      Top             =   510
      Width           =   795
   End
   Begin VB.TextBox txtUnsure 
      Alignment       =   1  'Right Justify
      Height          =   300
      Left            =   2805
      MaxLength       =   5
      TabIndex        =   1
      Text            =   "80"
      Top             =   120
      Width           =   795
   End
   Begin VB.Frame fraCodex 
      Height          =   30
      Left            =   75
      TabIndex        =   22
      Top             =   1005
      Width           =   5460
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   -30
      Top             =   7230
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagCodex.frx":058A
            Key             =   "ITEM"
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hgdCodex 
      Height          =   4020
      Left            =   90
      TabIndex        =   6
      Top             =   1395
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   7091
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483639
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      ScrollBars      =   2
      MergeCells      =   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   6
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   105
      Picture         =   "frmDiagCodex.frx":09DC
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblCodex 
      AutoSize        =   -1  'True
      Caption         =   "疾病辅助评估细则(&X):"
      Height          =   180
      Left            =   90
      TabIndex        =   5
      Top             =   1155
      Width           =   1800
   End
   Begin VB.Label lblEnsure 
      AutoSize        =   -1  'True
      Caption         =   "(&2) 当整体怀疑度达到          时，提示为临床诊断。"
      Height          =   180
      Left            =   960
      TabIndex        =   2
      Top             =   570
      Width           =   4500
   End
   Begin VB.Label lblUnsure 
      AutoSize        =   -1  'True
      Caption         =   "(&1) 当整体怀疑度达到          时，提示为疑似诊断；"
      Height          =   180
      Left            =   960
      TabIndex        =   0
      Top             =   180
      Width           =   4500
   End
End
Attribute VB_Name = "frmDiagCodex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mlngBarSize As Long

Dim rsTemp As New ADODB.Recordset
Dim objItem As ListItem
Dim strTemp As String
Dim intCount As Integer, lngRow As Integer, lngCol As Integer
Dim blnActive As Boolean

Const conCol分组名 As Integer = 0
Const conCol项目ID As Integer = 1
Const conCol项目名 As Integer = 2
Const conCol关系式 As Integer = 3
Const conCol条件值 As Integer = 4
Const conCol怀疑度 As Integer = 5

Private Sub cboFormula_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub cboGroup_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub cboValue_GotFocus()
    Me.cboValue.SelStart = 0: Me.cboValue.SelLength = 100
End Sub

Private Sub cboValue_KeyPress(KeyAscii As Integer)
    If InStr(" ~!@#$%^&*()_+|=`;'"":/<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub chkGroup_Click()
    If Not blnActive Then Exit Sub
    With Me.hgdCodex
        .Redraw = False
        .Clear
        .Rows = 1 + .FixedRows
        .TextMatrix(0, conCol分组名) = "分组名"
        .TextMatrix(0, conCol项目ID) = "项目ID"
        .TextMatrix(0, conCol项目名) = "项目名"
        .TextMatrix(0, conCol关系式) = "关系式"
        .TextMatrix(0, conCol条件值) = "条件值"
        .TextMatrix(0, conCol怀疑度) = "怀疑度"
        If Me.chkGroup.Value = 1 Then
            .ColWidth(conCol分组名) = 1000
            Me.txtGroup.Enabled = True
            Me.txtGroup.BackColor = &H80000005
        Else
            .ColWidth(conCol分组名) = 0
            Me.txtGroup.Text = ""
            Me.txtGroup.Enabled = False
            Me.txtGroup.BackColor = &H8000000F
        End If
        .ColWidth(conCol条件值) = .Width - .ColWidth(conCol分组名) - .ColWidth(conCol项目名) - .ColWidth(conCol关系式) - .ColWidth(conCol怀疑度) - mlngBarSize
        .Redraw = True
    End With
    MsgBox "由于评估方式改变，已经清除建立的全部评估细则！", vbExclamation, gstrSysName
    Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdAppend_Click()
    '细则定义正确性检查
    If Me.chkGroup.Value = 1 Then
        If Me.Tag = "西医" Then
            If Trim(Me.txtGroup.Text) = "" Then
                MsgBox "分组评估必须说明分组名！", vbExclamation, gstrSysName
                Me.txtGroup.SetFocus: Exit Sub
            End If
        Else
            If Trim(Me.cboGroup.Text) = "" Then
                MsgBox "中医必须说明分证候评估！" & vbCrLf & "如果无法选择证候，则请首先编辑参考并正确辨证！", vbExclamation, gstrSysName
                Me.cboGroup.SetFocus: Exit Sub
            End If
        End If
    End If
    If Trim(Me.txtItem.Tag) <> Trim(Me.txtItem.Text) Or Trim(Me.txtItem.Text) = "" Then
        MsgBox "未指定明确的评估细则项目！", vbExclamation, gstrSysName
        Me.txtItem.SetFocus: Exit Sub
    End If
    If Trim(Me.cboFormula.Text) = "" Then
        MsgBox "未指定明确的评估细则关系式！", vbExclamation, gstrSysName
        Me.cboFormula.SetFocus: Exit Sub
    End If
    If Me.cboValue.Enabled Then
        If Trim(Me.cboValue.Text) = "" Then
            MsgBox "未指定的评估细则条件值！", vbExclamation, gstrSysName
            Me.cboValue.SetFocus: Exit Sub
        End If
        strTemp = zlVerifyForm
        If strTemp <> "" Then
            MsgBox strTemp, vbExclamation, gstrSysName
            Me.cboValue.SetFocus: Exit Sub
        End If
    End If
    If Val(Me.txtDegree.Text) = 0 Then
        MsgBox "需要明确该细则在评估中的怀疑度，才能进行有效的评估！", vbExclamation, gstrSysName
        Me.txtDegree.SetFocus: Exit Sub
    End If
    
    '将细则添加到表格中：找到与当前细则组名相同的最后一行细则，再其后插入，找不到则最后插入
    Dim intAppendRow As Integer     '记录插入位置的行
    With Me.hgdCodex
        .Redraw = False
        If .Rows = 1 + .FixedRows And Val(.TextMatrix(.FixedRows, conCol项目ID)) = 0 Then
            intAppendRow = .FixedRows
        Else
            intAppendRow = .Rows
            For lngRow = .Rows - 1 To .FixedRows Step -1
                If Me.Tag = "西医" And .TextMatrix(lngRow, conCol分组名) = Trim(Me.txtGroup.Text) _
                   Or Me.Tag <> "西医" And .TextMatrix(lngRow, conCol分组名) = Trim(Me.cboGroup.Text) Then
                    intAppendRow = lngRow + 1: Exit For
                End If
            Next
            .Rows = .Rows + 1
            For lngRow = .Rows - 2 To intAppendRow Step -1
                For lngCol = 0 To .Cols - 1
                    .TextMatrix(lngRow + 1, lngCol) = .TextMatrix(lngRow, lngCol)
                Next
            Next
        End If
        .TextMatrix(intAppendRow, conCol分组名) = IIf(Me.Tag = "西医", Trim(Me.txtGroup.Text), Trim(Me.cboGroup.Text))
        .TextMatrix(intAppendRow, conCol项目ID) = Val(Me.lblItem.Tag)
        .TextMatrix(intAppendRow, conCol项目名) = Trim(Me.txtItem.Text)
        .TextMatrix(intAppendRow, conCol关系式) = Trim(Me.cboFormula.Text)
        .TextMatrix(intAppendRow, conCol条件值) = Trim(Me.cboValue.Text)
        .TextMatrix(intAppendRow, conCol怀疑度) = Val(Me.txtDegree.Text)
        .Row = intAppendRow
        .Col = conCol项目名
        .Redraw = True
    End With
    
    '清除细则定义控件内容，以便定义新的细则：
    Me.lblItem.Tag = ""
    Me.txtItem.Text = ""
    Me.txtItem.Tag = ""
    Me.lblValue.Tag = ""
    Me.cboValue.Text = ""
    Me.lblFormula.Tag = ""
    Me.cboFormula.Clear
    Me.hgdCodex.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdMove_Click(Index As Integer)
    With Me.hgdCodex
        If Index = 0 Then
            If .Row = .FixedRows Then Exit Sub
            If .TextMatrix(.Row, conCol分组名) <> .TextMatrix(.Row - 1, conCol分组名) Then
                MsgBox "评估细则不能在不同分组之间移动！", vbExclamation, gstrSysName: Exit Sub
            End If
        Else
            If .Row = .Rows - 1 Then Exit Sub
            If .TextMatrix(.Row, conCol分组名) <> .TextMatrix(.Row + 1, conCol分组名) Then
                MsgBox "评估细则不能在不同分组之间移动！", vbExclamation, gstrSysName: Exit Sub
            End If
        End If
        
        .Redraw = False
        strTemp = ""
        For lngCol = 0 To .Cols - 1
            strTemp = strTemp & "|" & .TextMatrix(.Row, lngCol)
        Next
        For lngCol = 0 To .Cols - 1
            If Index = 0 Then
                .TextMatrix(.Row, lngCol) = .TextMatrix(.Row - 1, lngCol)
                .TextMatrix(.Row - 1, lngCol) = Split(Mid(strTemp, 2), "|")(lngCol)
            Else
                .TextMatrix(.Row, lngCol) = .TextMatrix(.Row + 1, lngCol)
                .TextMatrix(.Row + 1, lngCol) = Split(Mid(strTemp, 2), "|")(lngCol)
            End If
        Next
        If Index = 0 Then
            .Row = .Row - 1
        Else
            .Row = .Row + 1
        End If
        .Redraw = True
    End With
    Call hgdCodex_RowColChange
End Sub

Private Sub cmdOK_Click()
    Dim intGrpNo As Integer, intItmNo As Integer
    If Val(Me.txtUnsure.Text) = 0 And Val(Me.txtEnsure.Text) = 0 Then
        MsgBox "为填写疑似诊断和临床诊断要求的怀疑度！", vbExclamation, gstrSysName
        Me.txtUnsure.SetFocus: Exit Sub
    End If
    If Val(Me.txtEnsure.Text) <> 0 And Val(Me.txtEnsure.Text) < Val(Me.txtUnsure.Text) Then
        MsgBox "疑似诊断怀疑度不应高于临床诊断要求的怀疑度！", vbExclamation, gstrSysName
        Me.txtUnsure.SetFocus: Exit Sub
    End If
    
    With Me.hgdCodex
        strTemp = "未分组"
        intGrpNo = -1: intItmNo = 0: gstrSql = ""
        For lngRow = .FixedRows To .Rows - 1
            If Trim(Me.hgdCodex.TextMatrix(Me.hgdCodex.FixedRows, conCol项目ID)) <> "" Then
                If strTemp <> Trim(.TextMatrix(lngRow, conCol分组名)) Then
                    intGrpNo = intGrpNo + 1: intItmNo = 0
                End If
                intItmNo = intItmNo + 1
                gstrSql = gstrSql & "|" & _
                    intGrpNo & "^" & Trim(.TextMatrix(lngRow, conCol分组名)) & "^" & _
                    intItmNo & "^" & Trim(.TextMatrix(lngRow, conCol项目ID)) & "^" & _
                    Trim(.TextMatrix(lngRow, conCol关系式)) & "^" & _
                    Trim(.TextMatrix(lngRow, conCol条件值)) & "^" & _
                    Trim(.TextMatrix(lngRow, conCol怀疑度))
                strTemp = Trim(.TextMatrix(lngRow, conCol分组名))
            End If
        Next
    End With
    If gstrSql = "" Then
        If MsgBox("未定义任何评估细则,如果确定，将删除给评估细则。" & vbCrLf & "继续吗？", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
            Me.hgdCodex.SetFocus: Exit Sub
        Else
            gstrSql = "zl_疾病诊断规则_Update(" & Me.hgdCodex.Tag & ",0,0,'')"
        End If
    Else
        gstrSql = "zl_疾病诊断规则_Update(" & _
                Me.hgdCodex.Tag & "," & _
                Val(Trim(Me.txtUnsure.Text)) & "," & _
                Val(Trim(Me.txtEnsure.Text)) & "," & _
                "'" & Mid(gstrSql, 2) & "')"
    End If
    Err = 0: On Error GoTo ErrHand
    Call zldatabase.ExecuteProcedure(gstrSql, Me.Caption)
    Unload Me
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdRemove_Click()
    If Val(Me.hgdCodex.TextMatrix(Me.hgdCodex.Row, conCol项目ID)) = 0 Then Exit Sub
    Err = 0: On Error GoTo ErrHand
    
    gstrSql = "select id, 编码, 中文名, 英文名, 类型, 长度, 小数, 单位, 表示法,数值域" & _
            " from 诊治所见项目 I" & _
            " where id=[1]"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Me.hgdCodex.TextMatrix(Me.hgdCodex.Row, conCol项目ID)))
    
    With rsTemp
        Me.lblItem.Tag = !ID
        Me.txtItem.Text = !中文名
        Me.txtItem.Tag = !中文名
        Me.lblFormula.Tag = IIf(IsNull(!类型), 0, !类型)
        Me.lblValue.Tag = IIf(IsNull(!数值域), "", !数值域)
        Me.cboValue.Tag = IIf(IsNull(!单位), "", !单位)
        Call zlAdjustForm
    End With
    
    Err = 0: On Error GoTo 0
    With Me.hgdCodex
        Me.cboFormula.Text = .TextMatrix(.Row, conCol关系式)
        Me.cboValue.Text = .TextMatrix(.Row, conCol条件值)
        Me.txtDegree.Text = .TextMatrix(.Row, conCol怀疑度)
        For lngRow = .Row To .Rows - 2
            For lngCol = 0 To .Cols - 1
                .TextMatrix(lngRow, lngCol) = .TextMatrix(lngRow + 1, lngCol)
            Next
        Next
        If .Rows = 1 + .FixedRows Then
            For lngCol = 0 To .Cols - 1
                .TextMatrix(.FixedRows, lngCol) = ""
            Next
        Else
            .Rows = .Rows - 1
        End If
    End With
    Call hgdCodex_RowColChange
    Me.txtItem.SetFocus
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
    If blnActive Then Exit Sub
    
    Err = 0: On Error GoTo ErrHand
    
    '评估总则填写
    gstrSql = "select ID,名称,疑似,临床" & _
            " from 疾病诊断目录" & _
            " where ID=[1]"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Me.hgdCodex.Tag))
    
    With rsTemp
        Me.Caption = !名称 & "．评估规则"
        Me.txtUnsure.Text = IIf(IsNull(!疑似), 0, !疑似)
        Me.txtEnsure.Text = IIf(IsNull(!临床), 0, !临床)
    End With
        
    '评估细则填写
    Me.hgdCodex.Redraw = False
    intCount = 0
    
    gstrSql = "select R.分组名,R.项目ID,I.中文名 as 项目名,R.关系式,R.条件值,R.怀疑度" & _
            " from 疾病诊断规则 R,诊治所见项目 I" & _
            " where R.项目ID=I.ID and R.诊断ID=[1] " & _
            " order by R.分组号,R.条件号"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Me.hgdCodex.Tag))
    
    With rsTemp
        If .EOF Then
            Me.hgdCodex.Rows = 1 + Me.hgdCodex.FixedRows
        Else
            Me.hgdCodex.Rows = .RecordCount + Me.hgdCodex.FixedRows
        End If
        Do While Not .EOF
            If Trim(IIf(IsNull(!分组名), "", !分组名)) <> "" Then
                intCount = 1
            End If
            Me.hgdCodex.TextMatrix(.AbsolutePosition, conCol分组名) = IIf(IsNull(!分组名), "", !分组名)
            Me.hgdCodex.TextMatrix(.AbsolutePosition, conCol项目ID) = IIf(IsNull(!项目ID), 0, !项目ID)
            Me.hgdCodex.TextMatrix(.AbsolutePosition, conCol项目名) = IIf(IsNull(!项目名), "", !项目名)
            Me.hgdCodex.TextMatrix(.AbsolutePosition, conCol关系式) = IIf(IsNull(!关系式), "", !关系式)
            Me.hgdCodex.TextMatrix(.AbsolutePosition, conCol条件值) = IIf(IsNull(!条件值), "", !条件值)
            Me.hgdCodex.TextMatrix(.AbsolutePosition, conCol怀疑度) = IIf(IsNull(!怀疑度), 0, !怀疑度)
            .MoveNext
        Loop
    End With
    
    If Me.Tag <> "西医" Then
        gstrSql = "select distinct 证候名称" & _
                " from 疾病诊断参考" & _
                " where 诊断id=[1] " & _
                "       and 证候名称 is not null"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Me.hgdCodex.Tag))
        
        Me.cboGroup.Clear
        Do While Not rsTemp.EOF
            Me.cboGroup.AddItem rsTemp!证候名称
            rsTemp.MoveNext
        Loop
    End If
    
    With Me.hgdCodex
        If Me.Tag = "西医" Then
            Me.txtGroup.Visible = True
            Me.cboGroup.Visible = False
            Me.lblNote.Caption = "(通常按疾病分型分期进行分组)"
            Me.chkGroup.Caption = "分组评估(&G)"
            Me.chkGroup.Enabled = True
            If intCount = 1 Then
                Me.chkGroup.Value = 1
                .ColWidth(conCol分组名) = 1000
                Me.txtGroup.Enabled = True
                Me.txtGroup.BackColor = &H80000005
                Me.lblNote.Visible = True
            Else
                Me.chkGroup.Value = 0
                .ColWidth(conCol分组名) = 0
                Me.txtGroup.Enabled = False
                Me.txtGroup.BackColor = &H8000000F
                Me.lblNote.Visible = False
            End If
        Else
            Me.txtGroup.Visible = False
            Me.cboGroup.Visible = True
            Me.lblNote.Caption = "(要求按参考中已经建立的辨证分类进行分组)"
            Me.lblNote.Visible = True
            Me.chkGroup.Caption = "辨证评估(&G)"
            Me.chkGroup.Value = 1
            Me.chkGroup.Enabled = False
            .ColWidth(conCol分组名) = 1000
        End If
        .ColWidth(conCol条件值) = .Width - .ColWidth(conCol分组名) - .ColWidth(conCol项目名) - .ColWidth(conCol关系式) - .ColWidth(conCol怀疑度) - mlngBarSize
        .Row = .FixedRows
        .Col = conCol项目名
    End With
    
    Me.hgdCodex.Redraw = True
    blnActive = True
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyEscape Then Exit Sub
    If Me.lvwList.Visible Then
        Me.lvwList.Visible = False
        Me.txtItem.SetFocus
    Else
        Call cmdCancel_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    blnActive = False
    With Me.hgdCodex
        .Redraw = False
        .ColAlignment(conCol分组名) = 0
        .ColAlignment(conCol项目ID) = 1
        .ColAlignment(conCol项目名) = 1
        .ColAlignment(conCol关系式) = 1
        .ColAlignment(conCol条件值) = 1
        .ColAlignment(conCol怀疑度) = 6
        
        .TextMatrix(0, conCol分组名) = "分组名"
        .TextMatrix(0, conCol项目ID) = "项目ID"
        .TextMatrix(0, conCol项目名) = "项目名"
        .TextMatrix(0, conCol关系式) = "关系式"
        .TextMatrix(0, conCol条件值) = "条件值"
        .TextMatrix(0, conCol怀疑度) = "怀疑度"
        .MergeCol(0) = True
        
        .ColWidth(conCol分组名) = 0
        .ColWidth(conCol项目ID) = 0
        .ColWidth(conCol项目名) = 1600
        .ColWidth(conCol关系式) = 900
        .ColWidth(conCol怀疑度) = 650
        .ColWidth(conCol条件值) = .Width - .ColWidth(conCol分组名) - .ColWidth(conCol项目名) - .ColWidth(conCol关系式) - .ColWidth(conCol怀疑度) - mlngBarSize
        .Redraw = True
    End With
    With Me.lvwList.ColumnHeaders
        .Clear
        .Add , "中文名", "中文名", 1800
        .Add , "编码", "编码", 1000
        .Add , "类型", "类型", 600
        .Add , "数值域", "数值域", 4000
    End With
    Me.lvwList.ColumnHeaders("编码").Position = 1
End Sub

Private Sub hgdCodex_GotFocus()
    Call hgdCodex_RowColChange
End Sub

Private Sub hgdCodex_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub hgdCodex_RowColChange()
    Dim lngCurRow As Long
    With Me.hgdCodex
        If Me.Tag = "西医" Then
            Me.txtGroup.Text = Left(.TextMatrix(.Row, conCol分组名), 10)
        Else
            For intCount = 0 To Me.cboGroup.ListCount - 1
                If Me.cboGroup.List(intCount) = .TextMatrix(.Row, conCol分组名) Then
                    Me.cboGroup.ListIndex = intCount
                End If
            Next
        End If
        .Redraw = False
        lngCurRow = .Row
        For lngRow = .FixedRows To .Rows - 1
            .Row = lngRow
            For lngCol = .FixedCols To .Cols - 1
                .Col = lngCol
                If lngRow = lngCurRow Then
                    .CellBackColor = &H80000001
                    .CellForeColor = &H80000005
                Else
                    .CellBackColor = .BackColor
                    .CellForeColor = .ForeColor
                End If
            Next
        Next
        .Row = lngCurRow
        .Col = conCol项目名
        .Redraw = True
    End With
End Sub

Private Sub lvwList_DblClick()
    If Me.lvwList.SelectedItem Is Nothing Then Exit Sub
    With Me.lvwList
        Me.lblItem.Tag = Mid(.SelectedItem.Key, 2)
        Me.txtItem.Text = Split(.SelectedItem.Tag, ",")(0)
        Me.txtItem.Tag = Split(.SelectedItem.Tag, ",")(0)
        Me.lblFormula.Tag = Split(.SelectedItem.Tag, ",")(1)
        Me.lblValue.Tag = .SelectedItem.SubItems(Me.lvwList.ColumnHeaders("数值域").Index - 1)
        Me.cboValue.Tag = Split(.SelectedItem.Tag, ",")(2)
        Call zlAdjustForm
        Me.txtItem.SetFocus
        Call zlCommFun.PressKey(vbKeyTab)
    End With
End Sub

Private Sub lvwList_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn, vbKeySpace
        If Me.lvwList.SelectedItem Is Nothing Then Exit Sub
        Call lvwList_DblClick
    End Select
End Sub

Private Sub lvwList_LostFocus()
    Me.lvwList.Visible = False
End Sub

Private Sub txtDegree_GotFocus()
    Me.txtDegree.SelStart = 0: Me.txtDegree.SelLength = 100
End Sub

Private Sub txtDegree_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txtEnsure_GotFocus()
    Me.txtEnsure.SelStart = 0: Me.txtEnsure.SelLength = 100
End Sub

Private Sub txtEnsure_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txtGroup_GotFocus()
    Call zlCommFun.OpenIme(True)
    Me.txtGroup.SelStart = 0: Me.txtGroup.SelLength = 100
End Sub

Private Sub txtGroup_KeyPress(KeyAscii As Integer)
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub txtGroup_LostFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txtItem_GotFocus()
    Me.txtItem.SelStart = 0: Me.txtItem.SelLength = 100
End Sub

Private Sub txtItem_KeyPress(KeyAscii As Integer)
    If InStr(" ~!@#$^&*()+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii <> vbKeyReturn Then Exit Sub
    Err = 0: On Error GoTo ErrHand
    
    '基础项目多数不能作为评估项目，但性别、年龄、职业例外
    gstrSql = "select I.id, I.编码, I.中文名, I.英文名, I.类型, I.长度, I.小数, I.单位, I.表示法,I.数值域" & _
            " from 诊治所见项目 I,诊治所见分类 C" & _
            " where I.分类ID=C.ID" & _
            "       And (C.性质=1 And C.编码 Not In ('01', '03', '04', '05') And I.中文名 <> '姓名' Or C.性质<>1)" & _
            "       And (I.编码 like [1] " & _
            "           or I.中文名 like [2] " & _
            "           or Upper(I.英文名) like [2])"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, Trim(Me.txtItem.Text) & "%", gstrMatch & Trim(Me.txtItem.Text) & "%")
    
    With rsTemp
        If .RecordCount = 0 Then
            MsgBox "未找到指定诊治所见项目", vbExclamation, gstrSysName
            Me.txtItem.SelStart = 0: Me.txtItem.SelLength = 100
            Me.txtItem.SetFocus
            Exit Sub
        End If
        If .RecordCount = 1 Then
            Me.lblItem.Tag = !ID
            Me.txtItem.Text = !中文名
            Me.txtItem.Tag = !中文名
            Me.lblFormula.Tag = IIf(IsNull(!类型), 0, !类型)
            Me.lblValue.Tag = IIf(IsNull(!数值域), "", !数值域)
            Me.cboValue.Tag = IIf(IsNull(!单位), "", !单位)
            Call zlAdjustForm
            KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab): Exit Sub
            Exit Sub
        End If
        
        Me.lvwList.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwList.ListItems.Add(, "_" & !ID, !中文名 & IIf(IsNull(!英文名), "", "(" & !英文名 & ")"), "ITEM", "ITEM")
            objItem.SubItems(Me.lvwList.ColumnHeaders("编码").Index - 1) = !编码
            Select Case IIf(IsNull(!类型), 0, !类型)
            Case 0
                objItem.SubItems(Me.lvwList.ColumnHeaders("类型").Index - 1) = "数值"
            Case 1
                objItem.SubItems(Me.lvwList.ColumnHeaders("类型").Index - 1) = "文字"
            Case 2
                objItem.SubItems(Me.lvwList.ColumnHeaders("类型").Index - 1) = "日期"
            Case 3
                objItem.SubItems(Me.lvwList.ColumnHeaders("类型").Index - 1) = "逻辑"
            End Select
            objItem.SubItems(Me.lvwList.ColumnHeaders("数值域").Index - 1) = IIf(IsNull(!数值域), "", !数值域)
            objItem.Tag = !中文名 & "," & IIf(IsNull(!类型), 0, !类型) & "," & IIf(IsNull(!单位), "", !单位)
            .MoveNext
        Loop
        With Me.lvwList
            .ListItems(1).Selected = True
            .Left = Me.fraDefine.Left + Me.txtItem.Left
            .Top = Me.fraDefine.Top + Me.txtItem.Top - .Height
            .Visible = True
            .SetFocus
        End With
    End With
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtUnsure_GotFocus()
    Me.txtUnsure.SelStart = 0: Me.txtUnsure.SelLength = 100
End Sub

Private Sub txtUnsure_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub zlAdjustForm()
    '-------------------------------------------------
    '调整条件表达式的可选范围
    '入参： 保存在Me.lblFormula.Tag中的数值类型，Me.lblValue.Tag中的数值域
    '-------------------------------------------------
    Dim aryValue() As String
    Me.cboValue.Clear
    Me.cboValue.Enabled = False
    Me.cboFormula.Clear
    Select Case Val(Me.lblFormula.Tag)
    Case 0  '数值
        If Me.cboValue.Tag = "" Then
            Me.lblValue.Caption = "值(&V):(数值型)"
        Else
            Me.lblValue.Caption = "值(&V):(数值型 单位:" & Me.cboValue.Tag & ")"
        End If
        Me.cboFormula.AddItem "等于"
        Me.cboFormula.AddItem "不等于"
        Me.cboFormula.AddItem "大于"
        Me.cboFormula.AddItem "小于"
        Me.cboFormula.AddItem "至多"
        Me.cboFormula.AddItem "至少"
        Me.cboFormula.AddItem "介于"
        Me.cboFormula.AddItem "存在"
        Me.cboFormula.AddItem "不存在"
        Me.cboFormula.ListIndex = 0
        Me.cboValue.Enabled = True
    Case 1  '文字
        Me.lblValue.Caption = "值(&V):(文字型)"
        Me.cboFormula.AddItem "等于"
        Me.cboFormula.AddItem "不等于"
        Me.cboFormula.AddItem "包含"
        Me.cboFormula.AddItem "不包含"
        Me.cboFormula.AddItem "存在"
        Me.cboFormula.AddItem "不存在"
        Me.cboFormula.ListIndex = 0
        Me.cboValue.Enabled = True
    Case 2  '日期
        Me.lblValue.Caption = "值(&V):(日期型)"
        Me.cboFormula.AddItem "等于"
        Me.cboFormula.AddItem "不等于"
        Me.cboFormula.AddItem "晚于"
        Me.cboFormula.AddItem "早于"
        Me.cboFormula.AddItem "不晚于"
        Me.cboFormula.AddItem "不早于"
        Me.cboFormula.AddItem "介于"
        Me.cboFormula.ListIndex = 0
        Me.cboValue.Enabled = True
    Case 3  '逻辑
        Me.lblValue.Caption = "值(&V):(逻辑型)"
        Me.cboFormula.AddItem "是"
        Me.cboFormula.AddItem "否"
        Me.cboFormula.ListIndex = 0
        Me.cboValue.Text = ""
        Me.cboValue.Enabled = False
    Case Else
    End Select
    
    aryValue = Split(Me.lblValue.Tag, ";")
    For intCount = LBound(aryValue) To UBound(aryValue)
        Me.cboValue.AddItem aryValue(intCount)
    Next
End Sub

Private Function zlVerifyForm() As String
    '-------------------------------------------------
    '判断条件表达式数值输入的正确性
    '入参：保存在Me.lblFormula.Tag中的数值类型
    '       Me.lblValue.Tag中的数值域，
    '       Me.lblFormula.text中的关系式
    '       Me.lblValue.text中的输入
    '出参：正确返回""，否则返回错误信息
    '-------------------------------------------------
    Dim aryValue() As String
    zlVerifyForm = ""
    On Error GoTo ErrHandle
    Select Case Val(Me.lblFormula.Tag)
    Case 0  '数值
        Select Case Me.cboFormula.Text
        Case "等于", "不等于", "大于", "小于", "至多", "至少"
            Me.cboValue.Text = Val(Me.cboValue.Text)
        Case "介于"
            aryValue = Split(Trim(Me.cboValue.Text), ",")
            If UBound(aryValue) <> 1 Then
                zlVerifyForm = "条件值未按“介于”要求规则“值1,值2”形式组织填写！": Exit Function
            End If
            Me.cboValue.Text = Val(aryValue(0)) & "," & Val(aryValue(1))
        Case "存在", "不存在"
            aryValue = Split(Trim(Me.cboValue.Text), ",")
            If UBound(aryValue) < 1 Then
                zlVerifyForm = "如果仅为单条件值，没必要采用“存在”或“不存在”的关系式！": Exit Function
            End If
            Me.cboValue.Text = ""
            For intCount = LBound(aryValue) To UBound(aryValue)
                Me.cboValue.Text = Me.cboValue.Text & "," & Val(aryValue(intCount))
            Next
            Me.cboValue.Text = Mid(Me.cboValue.Text, 2)
        End Select
    Case 1  '文字
        Select Case Me.cboFormula.Text
        Case "等于", "不等于", "包含", "不包含"
        Case "存在", "不存在"
            aryValue = Split(Trim(Me.cboValue.Text), ",")
            If UBound(aryValue) < 1 Then
                zlVerifyForm = "如果仅为单条件值，没必要采用“存在”或“不存在”的关系式！": Exit Function
            End If
        End Select
    Case 2  '日期
        Select Case Me.cboFormula.Text
        Case "等于", "不等于", "晚于", "早于", "不晚于", "不早于"
            gstrSql = "select to_char(to_date('" & Trim(Me.cboValue.Text) & "','YYYY-MM-DD'),'YYYY-MM-DD') from dual"
            With rsTemp
'                If .State = adStateOpen Then .Close
                Err = 0: On Error Resume Next
                Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "zlVerifyForm")
                If Err <> 0 Then zlVerifyForm = "输入条件值不符合日期格式规定(YYYY-MM-DD)！": Exit Function
                Err = 0: On Error GoTo 0
                Me.cboValue.Text = .Fields(0).Value
            End With
        Case "介于"
            aryValue = Split(Trim(Me.cboValue.Text), ",")
            If UBound(aryValue) <> 1 Then
                zlVerifyForm = "条件值未按“介于”要求规则“值1,值2”形式组织填写！": Exit Function
            End If
            Me.cboValue.Text = ""
            For intCount = LBound(aryValue) To UBound(aryValue)
                gstrSql = "select to_char(to_date('" & Trim(aryValue(intCount)) & "','YYYY-MM-DD'),'YYYY-MM-DD') from dual"
                With rsTemp
'                    If .State = adStateOpen Then .Close
                    Err = 0: On Error Resume Next
                    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "zlVerifyForm")
                    If Err <> 0 Then zlVerifyForm = "输入条件值中第" & intCount + 1 & "项不符合日期格式规定(YYYY-MM-DD)！": Exit Function
                    Err = 0: On Error GoTo 0
                    aryValue(intCount) = .Fields(0).Value
                End With
                Me.cboValue.Text = Me.cboValue.Text & "," & aryValue(intCount)
            Next
            Me.cboValue.Text = Mid(Me.cboValue.Text, 2)
        End Select
    Case 3  '逻辑
    Case Else
    End Select
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

