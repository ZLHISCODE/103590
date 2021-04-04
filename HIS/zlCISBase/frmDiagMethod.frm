VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.4#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmDiagMethod 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "诊疗措施选择"
   ClientHeight    =   3870
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6735
   Icon            =   "frmDiagMethod.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraLine 
      Height          =   45
      Left            =   0
      TabIndex        =   7
      Top             =   690
      Width           =   7530
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "全部清除(&R)"
      Height          =   350
      Left            =   1575
      Picture         =   "frmDiagMethod.frx":058A
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3375
      Width           =   1290
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4275
      TabIndex        =   2
      Top             =   3375
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5370
      TabIndex        =   3
      Top             =   3375
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   285
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3375
      Width           =   1290
   End
   Begin MSComctlLib.ListView lvwItems 
      Height          =   2715
      Left            =   390
      TabIndex        =   6
      Top             =   3885
      Visible         =   0   'False
      Width           =   5565
      _ExtentX        =   9816
      _ExtentY        =   4789
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
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
   Begin MSComctlLib.ImageList imgList 
      Left            =   5745
      Top             =   3960
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
            Picture         =   "frmDiagMethod.frx":06D4
            Key             =   "ItemUse"
         EndProperty
      EndProperty
   End
   Begin ZL9BillEdit.BillEdit msfMothed 
      Height          =   2355
      Left            =   285
      TabIndex        =   1
      Top             =   855
      Width           =   6165
      _ExtentX        =   10874
      _ExtentY        =   4154
      CellAlignment   =   9
      Text            =   ""
      TextMatrix0     =   ""
      MaxDate         =   2958465
      MinDate         =   -53688
      Value           =   36395
      Cols            =   2
      RowHeight0      =   315
      RowHeightMin    =   315
      ColWidth0       =   1005
      BackColor       =   -2147483643
      BackColorBkg    =   -2147483643
      BackColorSel    =   10249818
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      ForeColorSel    =   -2147483634
      GridColor       =   -2147483630
      ColAlignment0   =   9
      ListIndex       =   -1
      CellBackColor   =   -2147483643
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   180
      Picture         =   "frmDiagMethod.frx":0C6E
      Top             =   105
      Width           =   480
   End
   Begin VB.Label lblnote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "    根据当前参考的内容，选择指定在本院可行的诊疗项目措施，以便医生在疾病诊断治疗过程中，方便地依据疾病快速下达医嘱。"
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   855
      TabIndex        =   0
      Top             =   165
      Width           =   5580
   End
   Begin VB.Menu mnuSelect 
      Caption         =   "类别选择"
      Visible         =   0   'False
      Begin VB.Menu mnuKind 
         Caption         =   "西成药(&1)"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmDiagMethod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strMethod As String

Private strInputed As String
Dim rsTemp As New ADODB.Recordset
Dim objNode As Node
Dim objItem As ListItem
Dim intCount As Integer, strTemp As String

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdClear_Click()
    Me.msfMothed.ClearBill
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdOK_Click()
    strMethod = ""
    With Me.msfMothed
        For intCount = 1 To .Rows - 1
            If Trim(.TextMatrix(intCount, 1)) <> "" And .RowData(intCount) <> 0 Then
                If InStr(1, strMethod & ";", ";" & .RowData(intCount) & ";") > 0 Then
                    MsgBox intCount & "行，重复选择了“" & .TextMatrix(intCount, 1) & "”！", vbInformation, gstrSysName
                    .SetFocus: Exit Sub
                End If
                strMethod = strMethod & "," & .RowData(intCount)
            End If
        Next
    End With
    If strMethod = "" Then
        If MsgBox("没有选择或全部清除了诊疗措施，继续吗？", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
    End If
    If strMethod <> "" Then strMethod = Mid(strMethod, 2)
    Me.Hide
End Sub

Private Sub Form_Activate()
    If strMethod = "" Then Exit Sub
    
    '装入已经选择诊疗措施项目
    Err = 0: On Error GoTo ErrHand
    
    gstrSql = "select I.ID,I.编码,I.名称,I.标本部位 as 部位,I.计算单位 as 单位,K.名称 as 类别" & _
            " from 诊疗项目目录 I,诊疗项目类别 K" & _
            " where I.类别=K.编码 and I.ID in ([1])" & _
            " order by I.编码"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, strMethod)
    
    With rsTemp
        Me.msfMothed.ClearBill
        Do While Not .EOF
            If Me.msfMothed.Rows - 1 < .AbsolutePosition Then Me.msfMothed.Rows = Me.msfMothed.Rows + 1
            Me.msfMothed.RowData(.AbsolutePosition) = !ID
            Me.msfMothed.TextMatrix(.AbsolutePosition, 0) = .AbsolutePosition
            Me.msfMothed.TextMatrix(.AbsolutePosition, 1) = "[" & !编码 & "]" & !名称
            Me.msfMothed.TextMatrix(.AbsolutePosition, 2) = IIf(IsNull(!部位), "", !部位)
            Me.msfMothed.TextMatrix(.AbsolutePosition, 3) = IIf(IsNull(!单位), "", !单位)
            Me.msfMothed.TextMatrix(.AbsolutePosition, 4) = IIf(IsNull(!类别), "", !类别)
            .MoveNext
        Loop
    End With
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyEscape Then Exit Sub
    If Me.lvwItems.Visible Then
        Me.lvwItems.Visible = False
        Me.msfMothed.SetFocus
    Else
        Call cmdCancel_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    With Me.msfMothed
        .Active = True
        .MsfObj.FixedCols = 1: .Rows = 2: .Cols = 5
        .TextMatrix(0, 0) = "": .TextMatrix(0, 1) = "诊疗措施": .TextMatrix(0, 2) = "说明": .TextMatrix(0, 3) = "单位": .TextMatrix(0, 4) = "类别"
        .ColData(0) = 5: .ColData(1) = 1: .ColData(2) = 5: .ColData(3) = 5: .ColData(4) = 5
        .ColWidth(0) = 250: .ColWidth(1) = 3300: .ColWidth(2) = 1500: .ColWidth(3) = 500: .ColWidth(4) = 500
        .TextMatrix(1, 0) = "1"
        .PrimaryCol = 1: .LocateCol = 1
        .Row = 1: .Col = 1
    End With
    
    Me.lvwItems.ListItems.Clear
    With Me.lvwItems.ColumnHeaders
        .Clear
        .Add , "名称", "名称", 2500
        .Add , "编码", "编码", 1000
        .Add , "部位", "部位", 1200
        .Add , "单位", "单位", 550
        .Add , "类别", "类别", 550
    End With
    With Me.lvwItems
        .ColumnHeaders("编码").Position = 1
        .SortKey = .ColumnHeaders("编码").Index - 1
        .SortOrder = lvwAscending
    End With
End Sub

Private Sub lvwItems_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If Me.lvwItems.SortKey = ColumnHeader.Index - 1 Then
        Me.lvwItems.SortOrder = IIf(Me.lvwItems.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        Me.lvwItems.SortKey = ColumnHeader.Index - 1
        Me.lvwItems.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwItems_DblClick()
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    With Me.lvwItems
        Me.msfMothed.Text = "[" & .SelectedItem.SubItems(.ColumnHeaders("编码").Index - 1) & "]" & .SelectedItem.Text
        Me.msfMothed.RowData(Me.msfMothed.Row) = Mid(.SelectedItem.Key, 2)
        Me.msfMothed.TextMatrix(Me.msfMothed.Row, 1) = Me.msfMothed.Text
        Me.msfMothed.TextMatrix(Me.msfMothed.Row, 2) = .SelectedItem.SubItems(.ColumnHeaders("部位").Index - 1)
        Me.msfMothed.TextMatrix(Me.msfMothed.Row, 3) = .SelectedItem.SubItems(.ColumnHeaders("单位").Index - 1)
        Me.msfMothed.TextMatrix(Me.msfMothed.Row, 4) = .SelectedItem.SubItems(.ColumnHeaders("类别").Index - 1)
        Me.msfMothed.SetFocus
        Call zlCommFun.PressKey(vbKeyReturn)
    End With
End Sub

Private Sub lvwItems_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn, vbKeySpace
        If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
        Call lvwItems_DblClick
    End Select
End Sub

Private Sub lvwItems_LostFocus()
    Me.lvwItems.Visible = False
End Sub

Private Sub msfMothed_AfterAddRow(Row As Long)
    With Me.msfMothed
        For intCount = Row To .Rows - 1
            .TextMatrix(intCount, 0) = intCount
        Next
    End With
End Sub

Private Sub msfMothed_AfterDeleteRow()
    With Me.msfMothed
        For intCount = IIf(.Row <> 1, .Row - 1, .Row) To .Rows - 1
            .TextMatrix(intCount, 0) = intCount
        Next
    End With
End Sub

Private Sub msfMothed_CommandClick()
    Err = 0: On Error GoTo ErrHand
    gstrSql = "select I.ID,I.编码,I.名称,I.标本部位 as 部位,I.计算单位 as 单位,K.名称 as 类别" & _
            " from 诊疗项目目录 I,诊疗项目类别 K" & _
            " where I.类别=K.编码 and I.单独应用=1" & _
            "       and (I.撤档时间 is null or I.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))" & _
            " order by I.编码"
'        If .State = adStateOpen Then .Close
'        Call SQLTest(App.ProductName, Me.Caption, gstrSql)
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "msfMothed_CommandClick")
'        Call SQLTest
    With rsTemp
        If .BOF Or .EOF Then
            MsgBox "目前没有建立诊疗项目，无法设置！", vbExclamation, gstrSysName: Exit Sub
        End If
        Me.lvwItems.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !名称)
            objItem.Icon = "ItemUse": objItem.SmallIcon = "ItemUse"
            objItem.SubItems(Me.lvwItems.ColumnHeaders("编码").Index - 1) = !编码
            objItem.SubItems(Me.lvwItems.ColumnHeaders("部位").Index - 1) = IIf(IsNull(!部位), "", !部位)
            objItem.SubItems(Me.lvwItems.ColumnHeaders("单位").Index - 1) = IIf(IsNull(!单位), "", !单位)
            objItem.SubItems(Me.lvwItems.ColumnHeaders("类别").Index - 1) = IIf(IsNull(!类别), "", !类别)
            .MoveNext
        Loop
        Me.lvwItems.ListItems(1).Selected = True
    End With
    With Me.lvwItems
        .Left = Me.msfMothed.Left + 300
        .Top = Me.msfMothed.Top
        .ZOrder 0: .Visible = True
        .SetFocus
    End With
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub msfMothed_EnterCell(Row As Long, Col As Long)
    strInputed = Me.msfMothed.TextMatrix(Row, Col)
End Sub

Private Sub msfMothed_GotFocus()
    If Me.lvwItems.Visible Then Me.lvwItems.SetFocus
End Sub

Private Sub msfMothed_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    If KeyCode <> vbKeyReturn Then Exit Sub
    With Me.msfMothed
        If .Active = False Then Exit Sub
        If .Col <> 1 Then Exit Sub
        If .TxtVisible = False Then
            If .TextMatrix(.Row, 1) = "" Then Exit Sub
            strTemp = UCase(Trim(.TextMatrix(.Row, 1)))
        Else
            If Trim(.Text) = "" Then Exit Sub
            strTemp = UCase(Trim(.Text))
        End If
    End With
    If strInputed = strTemp Then Exit Sub
    If Me.msfMothed.RowData(Me.msfMothed.Row) <> 0 And InStr(1, strTemp, "[") <> 0 And InStr(1, strTemp, "]") <> 0 Then
        gstrSql = "select I.ID,I.编码,I.名称,I.标本部位 as 部位,I.计算单位 as 单位,K.名称 as 类别" & _
                " from 诊疗项目目录 I,诊疗项目类别 K" & _
                " where I.类别=K.编码 and I.单独应用=1 and I.ID=[1] "
    Else
        If InStr(1, strTemp, "[") <> 0 And InStr(1, strTemp, "]") <> 0 Then strTemp = Mid(strTemp, 2, InStr(1, strTemp, "]") - 2)
        gstrSql = "select distinct I.ID,I.编码,I.名称,I.标本部位 as 部位,I.计算单位 as 单位,K.名称 as 类别" & _
                " from 诊疗项目目录 I,诊疗项目别名 N,诊疗项目类别 K" & _
                " where I.ID=N.诊疗项目id and I.类别=K.编码 and I.单独应用=1" & _
                "       and (I.撤档时间 is null or I.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))" & _
                "       and (I.编码 like [2]" & _
                "           or N.名称 like [3] " & _
                "           or N.简码 like [3])"
    End If
    
    Err = 0: On Error GoTo ErrHand
    
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, Me.msfMothed.RowData(Me.msfMothed.Row), strTemp & "%", gstrMatch & strTemp & "%")
        
    With rsTemp
        If .BOF Or .EOF Then
            MsgBox "未找到指定的诊疗项目，请重新输入！", vbExclamation, gstrSysName: Cancel = True: Exit Sub
        End If
        If .RecordCount = 1 Then
            Me.msfMothed.Text = "[" & !编码 & "]" & !名称
            Me.msfMothed.RowData(Me.msfMothed.Row) = !ID
            Me.msfMothed.TextMatrix(Me.msfMothed.Row, 1) = Me.msfMothed.Text
            Me.msfMothed.TextMatrix(Me.msfMothed.Row, 2) = IIf(IsNull(!部位), "", !部位)
            Me.msfMothed.TextMatrix(Me.msfMothed.Row, 3) = IIf(IsNull(!单位), "", !单位)
            Me.msfMothed.TextMatrix(Me.msfMothed.Row, 4) = IIf(IsNull(!类别), "", !类别)
            Exit Sub
        End If
        Me.lvwItems.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !名称)
            objItem.Icon = "ItemUse": objItem.SmallIcon = "ItemUse"
            objItem.SubItems(Me.lvwItems.ColumnHeaders("编码").Index - 1) = !编码
            objItem.SubItems(Me.lvwItems.ColumnHeaders("部位").Index - 1) = IIf(IsNull(!部位), "", !部位)
            objItem.SubItems(Me.lvwItems.ColumnHeaders("单位").Index - 1) = IIf(IsNull(!单位), "", !单位)
            objItem.SubItems(Me.lvwItems.ColumnHeaders("类别").Index - 1) = IIf(IsNull(!类别), "", !类别)
            .MoveNext
        Loop
        Me.lvwItems.ListItems(1).Selected = True
    End With
    With Me.lvwItems
        .Tag = Me.msfMothed.Name
        .Left = Me.msfMothed.Left + 300
        .Top = Me.msfMothed.Top
        .ZOrder 0: .Visible = True
        .SetFocus
    End With
    Cancel = True
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
