VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.4#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmMediTabu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "药品配伍禁忌"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6810
   Icon            =   "frmMediTabu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   6810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdMedi 
      Caption         =   "…"
      Height          =   285
      Left            =   6240
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   758
      Width           =   285
   End
   Begin VB.CommandButton cmdRestore 
      Caption         =   "恢复(&R)"
      Height          =   350
      Left            =   2715
      Picture         =   "frmMediTabu.frx":058A
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   4275
      Width           =   1290
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "全部清除(&C)"
      Height          =   350
      Left            =   1425
      Picture         =   "frmMediTabu.frx":06D4
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4275
      Width           =   1290
   End
   Begin VB.TextBox txtMedi 
      Height          =   300
      Left            =   1275
      MaxLength       =   50
      TabIndex        =   2
      Top             =   750
      Width           =   4980
   End
   Begin MSComctlLib.ListView lvwItems 
      Height          =   2790
      Left            =   465
      TabIndex        =   8
      Top             =   5100
      Visible         =   0   'False
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   4921
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
      Left            =   3030
      Top             =   4560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediTabu.frx":081E
            Key             =   "ItemUse"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediTabu.frx":0DB8
            Key             =   "ItemStop"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "保存(&S)"
      Height          =   350
      Left            =   4335
      TabIndex        =   5
      Top             =   4275
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   225
      Picture         =   "frmMediTabu.frx":1352
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4275
      Width           =   1100
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "关闭(&X)"
      Height          =   350
      Left            =   5445
      TabIndex        =   6
      Top             =   4275
      Width           =   1100
   End
   Begin ZL9BillEdit.BillEdit msfTabu 
      Height          =   2775
      Left            =   225
      TabIndex        =   4
      Top             =   1395
      Width           =   6285
      _ExtentX        =   11086
      _ExtentY        =   4895
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
      Left            =   240
      Picture         =   "frmMediTabu.frx":149C
      Top             =   0
      Width           =   480
   End
   Begin VB.Label lblTabu 
      AutoSize        =   -1  'True
      Caption         =   "禁忌药品(&T)"
      Height          =   180
      Left            =   240
      TabIndex        =   3
      Top             =   1155
      Width           =   990
   End
   Begin VB.Label lblMedi 
      AutoSize        =   -1  'True
      Caption         =   "指定药品(&M)"
      Height          =   180
      Left            =   255
      TabIndex        =   1
      Top             =   810
      Width           =   990
   End
   Begin VB.Label lblnote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "    请选择药品后，指定其配伍禁忌药品。本系统将禁忌分为慎用与禁用，在处方时将根据设置进行提醒或禁止提示。"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   870
      TabIndex        =   0
      Top             =   120
      Width           =   5685
   End
End
Attribute VB_Name = "frmMediTabu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------
'说明：
'   1、当前材质：由me.tag保存,分别为"5","6","7"
'   2、当前状态：由me.cmdClose.tag保存，分别为"修改"、"查阅"，由上级程序传入
'   3、指定药品：由me.lblMedi.tag保存，由上级程序传入可以传递，也可以不传递
'---------------------------------------------------
Private strInputed As String
Dim rsTemp As New ADODB.Recordset
Dim objItem As ListItem
Dim strTemp As String
Dim intCount As Integer

Private Sub cmdClear_Click()
    With Me.msfTabu
        .ClearBill
        .AddItem "慎用": .ItemData(.NewIndex) = 1
        .AddItem "禁用": .ItemData(.NewIndex) = 2
        .ListIndex = 0
    End With
End Sub

Private Sub cmdClose_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub cmdRestore_Click()
    Call zlTabuRef(Me.lblMedi.Tag)
End Sub

Private Sub cmdSave_Click()
    If Val(Me.lblMedi.Tag) = 0 Then MsgBox "未正确指定药品！", vbExclamation, gstrSysName: Me.txtMedi.SetFocus: Exit Sub
    '禁忌检查
    strTemp = "": gstrSql = ""
    With Me.msfTabu
        For intCount = 1 To .Rows - 1
            If Trim(.TextMatrix(intCount, 1)) <> "" And .RowData(intCount) <> 0 Then
                If InStr(1, strTemp & ";", ";" & .RowData(intCount) & ";") > 0 Then
                    MsgBox intCount & "行禁忌药品与前面的药品有重复！", vbInformation, gstrSysName
                    .SetFocus: Exit Sub
                End If
                strTemp = strTemp & ";" & .RowData(intCount)
                gstrSql = gstrSql & "|" & .RowData(intCount) & "^" & IIf(Trim(.TextMatrix(intCount, 2)) = "慎用", 1, 2)
            End If
        Next
    End With
    If gstrSql <> "" Then gstrSql = Mid(gstrSql, 2)
    gstrSql = "zl_配伍禁忌_UPDATE(" & Val(Me.lblMedi.Tag) & ",'" & gstrSql & "')"
    
    Err = 0: On Error GoTo ErrHand
    Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
    MsgBox Me.txtMedi.Text & " 相关配伍禁忌保存成功！", vbExclamation, gstrSysName
    Me.txtMedi.SetFocus
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdMedi_Click()
    Err = 0: On Error GoTo ErrHand
    
    gstrSql = "select I.ID,I.编码,I.名称,I.计算单位,T.药品剂型,T.毒理分类" & _
            " from 诊疗项目目录 I,药品特性 T" & _
            " where I.ID=T.药名ID and I.类别=[1] " & _
            "       and (I.撤档时间 is null or I.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Me.Tag)
    
    With rsTemp
        If .BOF Or .EOF Then
            MsgBox "请建立药品品种后设置配伍禁忌", vbExclamation, gstrSysName
            Me.lblMedi.Tag = 0: Me.txtMedi.Tag = "": Me.txtMedi.Text = Me.txtMedi.Tag: Me.txtMedi.SetFocus: Exit Sub
        End If
        If .RecordCount = 1 Then
            If Me.lblMedi.Tag <> !ID Then
                Me.lblMedi.Tag = !ID: Me.txtMedi.Tag = "[" & !编码 & "]" & !名称: Me.txtMedi.Text = Me.txtMedi.Tag
                Call zlTabuRef(Me.lblMedi.Tag)
            End If
            Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        Me.lvwItems.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !名称)
            objItem.Icon = "ItemUse": objItem.SmallIcon = "ItemUse"
            objItem.SubItems(Me.lvwItems.ColumnHeaders("编码").Index - 1) = !编码
            objItem.SubItems(Me.lvwItems.ColumnHeaders("计算单位").Index - 1) = IIf(IsNull(!计算单位), "", !计算单位)
            objItem.SubItems(Me.lvwItems.ColumnHeaders("药品剂型").Index - 1) = IIf(IsNull(!药品剂型), "", !药品剂型)
            objItem.SubItems(Me.lvwItems.ColumnHeaders("毒理分类").Index - 1) = !毒理分类
            .MoveNext
        Loop
        Me.lvwItems.ListItems(1).Selected = True
    End With
    With Me.lvwItems
        .Tag = Me.txtMedi.Name
        .Left = Me.txtMedi.Left
        .Top = Me.txtMedi.Top + Me.txtMedi.Height
        .ZOrder 0: .Visible = True
        .SetFocus
    End With
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
    If Me.cmdClose.Tag = "查阅" Then
        Me.msfTabu.Active = False
        Me.cmdSave.Visible = False
        Me.cmdClear.Visible = False
        Me.cmdRestore.Visible = False
    Else
        Me.msfTabu.Active = True
    End If
    
    Err = 0: On Error GoTo ErrHand
    
    gstrSql = "select I.ID,I.编码,I.名称,I.计算单位,T.药品剂型,T.毒理分类" & _
            " from 诊疗项目目录 I,药品特性 T" & _
            " where I.ID=T.药名ID  and I.类别=[1] and I.ID=[2] " & _
            "       and (I.撤档时间 is null or I.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Me.Tag, Val(Me.lblMedi.Tag))
    
    With rsTemp
        If .BOF Or .EOF Then
            Me.lblMedi.Tag = 0: Me.txtMedi.Tag = "": Me.txtMedi.Text = Me.txtMedi.Tag
        Else
            Me.lblMedi.Tag = !ID: Me.txtMedi.Tag = "[" & !编码 & "]" & !名称: Me.txtMedi.Text = Me.txtMedi.Tag
            Call zlTabuRef(Me.lblMedi.Tag)
        End If
    End With
    Me.txtMedi.SetFocus
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyEscape Then Exit Sub
    If Me.lvwItems.Visible Then
        Me.lvwItems.Visible = False
        If Me.lvwItems.Tag = Me.txtMedi.Name Then
            Me.txtMedi.SetFocus
        Else
            Me.msfTabu.SetFocus
        End If
    Else
        cmdClose_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    With Me.msfTabu
        .Active = True
        .MsfObj.FixedCols = 1: .Rows = 2: .Cols = 3
        .TextMatrix(0, 0) = "": .TextMatrix(0, 1) = "通用名称": .TextMatrix(0, 2) = "禁忌类型"
        .ColData(0) = 5: .ColData(1) = 1: .ColData(2) = 3
        .ColWidth(0) = 250: .ColWidth(1) = 3500: .ColWidth(2) = 1000
        .TextMatrix(1, 0) = "1"
        .PrimaryCol = 1: .LocateCol = 1
        .Row = 1: .Col = 1
        .AddItem "慎用": .ItemData(.NewIndex) = 1
        .AddItem "禁用": .ItemData(.NewIndex) = 2
        .ListIndex = 0
    End With
    
    Me.lvwItems.ListItems.Clear
    With Me.lvwItems.ColumnHeaders
        .Clear
        .Add , "名称", "名称", 2500
        .Add , "编码", "编码", 1000
        .Add , "计算单位", "剂量单位", 900
        .Add , "药品剂型", "剂型", 600
        .Add , "毒理分类", "毒理", 750
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
        If .Tag = Me.txtMedi.Name Then
            If Me.lblMedi.Tag <> Mid(.SelectedItem.Key, 2) Then
                Me.lblMedi.Tag = Mid(.SelectedItem.Key, 2)
                Me.txtMedi.Tag = "[" & .SelectedItem.SubItems(.ColumnHeaders("编码").Index - 1) & "]" & .SelectedItem.Text
                Me.txtMedi.Text = Me.txtMedi.Tag
                Call zlTabuRef(Me.lblMedi.Tag)
            End If
            Me.txtMedi.SetFocus
            Call zlCommFun.PressKey(vbKeyTab)
        Else
            Me.msfTabu.Text = "[" & .SelectedItem.SubItems(.ColumnHeaders("编码").Index - 1) & "]" & .SelectedItem.Text
            Me.msfTabu.TextMatrix(Me.msfTabu.Row, 1) = Me.msfTabu.Text
            Me.msfTabu.RowData(Me.msfTabu.Row) = Mid(.SelectedItem.Key, 2)
            Me.msfTabu.SetFocus
            Call zlCommFun.PressKey(vbKeyRight)
        End If
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

Private Sub msfTabu_AfterAddRow(Row As Long)
    With Me.msfTabu
        For intCount = Row To .Rows - 1
            .TextMatrix(intCount, 0) = intCount
        Next
    End With
End Sub

Private Sub msfTabu_AfterDeleteRow()
    With Me.msfTabu
        For intCount = IIf(.Row <> 1, .Row - 1, .Row) To .Rows - 1
            .TextMatrix(intCount, 0) = intCount
        Next
    End With
End Sub

Private Sub msfTabu_cboClick(ListIndex As Long)
    Me.msfTabu.TextMatrix(Me.msfTabu.Row, 2) = Me.msfTabu.CboText
End Sub

Private Sub msfTabu_CommandClick()
    Err = 0: On Error GoTo ErrHand
    
    gstrSql = "select I.ID,I.编码,I.名称,I.计算单位,T.药品剂型,T.毒理分类" & _
            " from 诊疗项目目录 I,药品特性 T" & _
            " where I.ID=T.药名ID and I.类别=[1] " & _
            "       and (I.撤档时间 is null or I.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))" & _
            "       and I.ID<>[2] "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Me.Tag, Me.lblMedi.Tag)
    
    With rsTemp
        If .BOF Or .EOF Then
            MsgBox "在没有建立其他药品品种时，无法设置配伍禁忌！", vbExclamation, gstrSysName: Exit Sub
        End If
        Me.lvwItems.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !名称)
            objItem.Icon = "ItemUse": objItem.SmallIcon = "ItemUse"
            objItem.SubItems(Me.lvwItems.ColumnHeaders("编码").Index - 1) = !编码
            objItem.SubItems(Me.lvwItems.ColumnHeaders("计算单位").Index - 1) = IIf(IsNull(!计算单位), "", !计算单位)
            objItem.SubItems(Me.lvwItems.ColumnHeaders("药品剂型").Index - 1) = IIf(IsNull(!药品剂型), "", !药品剂型)
            objItem.SubItems(Me.lvwItems.ColumnHeaders("毒理分类").Index - 1) = !毒理分类
            .MoveNext
        Loop
        Me.lvwItems.ListItems(1).Selected = True
    End With
    With Me.lvwItems
        .Tag = Me.msfTabu.Name
        .Left = Me.msfTabu.Left + 300
        .Top = Me.msfTabu.Top + Me.msfTabu.RowHeight(0)
        .ZOrder 0: .Visible = True
        .SetFocus
    End With
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub msfTabu_EnterCell(Row As Long, Col As Long)
    strInputed = Me.msfTabu.TextMatrix(Row, Col)
End Sub

Private Sub msfTabu_GotFocus()
    If Me.lvwItems.Visible Then Me.lvwItems.SetFocus
End Sub

Private Sub msfTabu_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    If KeyCode <> vbKeyReturn Then Exit Sub
    With Me.msfTabu
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
    If InStr(1, strTemp, "[") <> 0 And InStr(1, strTemp, "]") <> 0 Then strTemp = Mid(strTemp, 2, InStr(1, strTemp, "]") - 2)
    
    Err = 0: On Error GoTo ErrHand
    
    gstrSql = "select distinct I.ID,I.编码,I.名称,I.计算单位,T.药品剂型,T.毒理分类" & _
            " from 诊疗项目目录 I,药品特性 T,诊疗项目别名 N" & _
            " where I.ID=T.药名ID and I.ID=N.诊疗项目ID and I.类别=[1] " & _
            "       and (I.撤档时间 is null or I.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))" & _
            "       and (I.编码 like [2] or N.名称 like [3] or N.简码 like [3])" & _
            "       and I.ID<>[4] "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Me.Tag, strTemp & "%", gstrMatch & strTemp & "%", Val(Me.txtMedi.Tag))
    
    With rsTemp
        If .BOF Or .EOF Then
            MsgBox "未找到指定药品，请重新输入！", vbExclamation, gstrSysName: Cancel = True: Exit Sub
        End If
        If .RecordCount = 1 Then
            Me.msfTabu.Text = "[" & !编码 & "]" & !名称
            Me.msfTabu.TextMatrix(Me.msfTabu.Row, 1) = Me.msfTabu.Text
            Me.msfTabu.RowData(Me.msfTabu.Row) = !ID
            Exit Sub
        End If
        Me.lvwItems.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !名称)
            objItem.Icon = "ItemUse": objItem.SmallIcon = "ItemUse"
            objItem.SubItems(Me.lvwItems.ColumnHeaders("编码").Index - 1) = !编码
            objItem.SubItems(Me.lvwItems.ColumnHeaders("计算单位").Index - 1) = IIf(IsNull(!计算单位), "", !计算单位)
            objItem.SubItems(Me.lvwItems.ColumnHeaders("药品剂型").Index - 1) = IIf(IsNull(!药品剂型), "", !药品剂型)
            objItem.SubItems(Me.lvwItems.ColumnHeaders("毒理分类").Index - 1) = !毒理分类
            .MoveNext
        Loop
        Me.lvwItems.ListItems(1).Selected = True
    End With
    With Me.lvwItems
        .Tag = Me.msfTabu.Name
        .Left = Me.msfTabu.Left + 300
        .Top = Me.msfTabu.Top + Me.msfTabu.RowHeight(0)
        .ZOrder 0: .Visible = True
        .SetFocus
    End With
    Cancel = True
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtMedi_GotFocus()
    Me.txtMedi.SelStart = 0: Me.txtMedi.SelLength = 100
End Sub

Private Sub txtMedi_KeyPress(KeyAscii As Integer)
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If KeyAscii <> vbKeyReturn Then Exit Sub
    strTemp = UCase(Trim(Me.txtMedi.Text))
    If strTemp = "" Then Me.lblMedi.Tag = 0: Me.txtMedi.Tag = "": Me.txtMedi.Text = "": Exit Sub
    
    If InStr(1, strTemp, "[") <> 0 And InStr(1, strTemp, "]") <> 0 Then strTemp = Mid(strTemp, 2, InStr(1, strTemp, "]") - 2)
    Err = 0: On Error GoTo ErrHand
    
    gstrSql = "select distinct I.ID,I.编码,I.名称,I.计算单位,T.药品剂型,T.毒理分类" & _
            " from 诊疗项目目录 I,药品特性 T,诊疗项目别名 N" & _
            " where I.ID=T.药名ID and I.ID=N.诊疗项目ID and I.类别=[1] " & _
            "       and (I.撤档时间 is null or I.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))" & _
            "       and (I.编码 like [2] or N.名称 like [3] or N.简码 like [3])"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Me.Tag, strTemp & "%", gstrMatch & strTemp & "%")
    
    With rsTemp
        If .BOF Or .EOF Then
            MsgBox "未找到指定的药品，请重新指定", vbExclamation, gstrSysName
            Me.lblMedi.Tag = 0: Me.txtMedi.Tag = "": Me.txtMedi.Text = Me.txtMedi.Tag: Me.txtMedi.SetFocus: Exit Sub
        End If
        If .RecordCount = 1 Then
            If Me.lblMedi.Tag <> !ID Then
                Me.lblMedi.Tag = !ID: Me.txtMedi.Tag = "[" & !编码 & "]" & !名称: Me.txtMedi.Text = Me.txtMedi.Tag
                Call zlTabuRef(Me.lblMedi.Tag)
            End If
            Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        Me.lvwItems.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !名称)
            objItem.Icon = "ItemUse": objItem.SmallIcon = "ItemUse"
            objItem.SubItems(Me.lvwItems.ColumnHeaders("编码").Index - 1) = !编码
            objItem.SubItems(Me.lvwItems.ColumnHeaders("计算单位").Index - 1) = IIf(IsNull(!计算单位), "", !计算单位)
            objItem.SubItems(Me.lvwItems.ColumnHeaders("药品剂型").Index - 1) = IIf(IsNull(!药品剂型), "", !药品剂型)
            objItem.SubItems(Me.lvwItems.ColumnHeaders("毒理分类").Index - 1) = !毒理分类
            .MoveNext
        Loop
        Me.lvwItems.ListItems(1).Selected = True
    End With
    With Me.lvwItems
        .Tag = Me.txtMedi.Name
        .Left = Me.txtMedi.Left
        .Top = Me.txtMedi.Top + Me.txtMedi.Height
        .ZOrder 0: .Visible = True
        .SetFocus
    End With
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtMedi_LostFocus()
    Me.txtMedi.Text = Me.txtMedi.Tag
End Sub

Private Sub zlTabuRef(lngMediId As Long)
    '--------------------------------------------------------
    '功能：刷新显示药品的配伍禁忌
    '入参：lngMediId-指定的药名id
    '--------------------------------------------------------
    Err = 0: On Error GoTo ErrHand
    
    '配伍禁忌
    gstrSql = "select I.ID,I.名称,R.类型" & _
            " from 诊疗项目目录 I,诊疗互斥项目 R" & _
            " where I.ID=R.项目ID  and I.ID<>[1] " & _
            "       and R.组编号 in (select 组编号 from 诊疗互斥项目 where 项目ID=[1])"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngMediId)
    
    With rsTemp
        Me.msfTabu.ClearBill
        Do While Not .EOF
            If Me.msfTabu.Rows - 1 < .AbsolutePosition Then Me.msfTabu.Rows = Me.msfTabu.Rows + 1
            Me.msfTabu.TextMatrix(.AbsolutePosition, 0) = .AbsolutePosition
            Me.msfTabu.RowData(.AbsolutePosition) = !ID
            Me.msfTabu.TextMatrix(.AbsolutePosition, 1) = !名称
            Me.msfTabu.TextMatrix(.AbsolutePosition, 2) = IIf(!类型 = 1, "慎用", "禁用")
            .MoveNext
        Loop
    End With
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
