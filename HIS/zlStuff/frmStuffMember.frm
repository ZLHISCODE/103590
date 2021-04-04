VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.4#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmStuffMember 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "材料组成"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7875
   Icon            =   "frmStuffMember.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   7875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdMedi 
      Caption         =   "…"
      Height          =   285
      Left            =   7440
      TabIndex        =   12
      TabStop         =   0   'False
      Tag             =   "分类"
      ToolTipText     =   "按*打开选择器"
      Top             =   788
      Width           =   285
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "关闭(&X)"
      Height          =   350
      Left            =   6625
      TabIndex        =   5
      Top             =   4185
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   90
      Picture         =   "frmStuffMember.frx":058A
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   4185
      Width           =   1100
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "保存(&S)"
      Height          =   350
      Left            =   5535
      TabIndex        =   4
      Top             =   4185
      Width           =   1100
   End
   Begin VB.TextBox txtMedi 
      Height          =   300
      Left            =   1125
      MaxLength       =   50
      TabIndex        =   1
      Top             =   780
      Width           =   6285
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "全部清除(&C)"
      Height          =   350
      Left            =   1275
      Picture         =   "frmStuffMember.frx":06D4
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4185
      Width           =   1290
   End
   Begin VB.CommandButton cmdRestore 
      Caption         =   "恢复(&R)"
      Height          =   350
      Left            =   2565
      Picture         =   "frmStuffMember.frx":081E
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4185
      Width           =   1290
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   2880
      Top             =   4590
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
            Picture         =   "frmStuffMember.frx":0968
            Key             =   "ItemUse"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMember.frx":0F02
            Key             =   "ItemStop"
         EndProperty
      EndProperty
   End
   Begin ZL9BillEdit.BillEdit msfMember 
      Height          =   2430
      Left            =   90
      TabIndex        =   3
      Top             =   1620
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   4286
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
   Begin MSComctlLib.ListView lvwItems 
      Height          =   2505
      Left            =   960
      TabIndex        =   9
      Top             =   4710
      Visible         =   0   'False
      Width           =   6600
      _ExtentX        =   11642
      _ExtentY        =   4419
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
   Begin VB.Label lblSpec 
      AutoSize        =   -1  'True
      Caption         =   "规格：      厂牌：       单位：瓶"
      Height          =   180
      Left            =   1125
      TabIndex        =   11
      Top             =   1125
      Width           =   2970
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   75
      Picture         =   "frmStuffMember.frx":149C
      Top             =   90
      Width           =   480
   End
   Begin VB.Label lblnote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "    选择具体规格的的卫生材料，按散装单位指定其具体的组成卫生材料；未指定其组成（或清除其组成）"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   720
      TabIndex        =   10
      Top             =   210
      Width           =   7065
   End
   Begin VB.Label lblMedi 
      AutoSize        =   -1  'True
      Caption         =   "指定卫材(&M)"
      Height          =   180
      Left            =   90
      TabIndex        =   0
      Top             =   840
      Width           =   990
   End
   Begin VB.Label lblMember 
      AutoSize        =   -1  'True
      Caption         =   "组成的卫生材料(&E)："
      Height          =   180
      Left            =   90
      TabIndex        =   2
      Top             =   1395
      Width           =   1710
   End
End
Attribute VB_Name = "frmStuffMember"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------
'说明：
'   1、当前材质：卫生材料
'   2、当前状态：由me.cmdClose.tag保存，分别为"修改"、"查阅"，由上级程序传入
'   3、指定卫材：由me.lblMedi.tag保存，由上级程序传入可以传递，也可以不传递
'   4、当前编辑内容：由Me.msfMember.Tag保存，分别为"自制"
Dim mobjItem As ListItem
Dim mstrTemp  As String
Dim mintCount As Integer

Private Const col品名 As Integer = 1
Private Const col规格 As Integer = 2
Private Const col产地 As Integer = 3
Private Const col采用量 As Integer = 4
Private Const col单位 As Integer = 5

Private Sub cmdClear_Click()
    Me.msfMember.ClearBill
End Sub

Private Sub cmdClose_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub cmdRestore_Click()
    Call zlMemberRef(Me.lblMedi.Tag)
End Sub

Private Sub CmdSave_Click()
    If Val(Me.lblMedi.Tag) = 0 Then ShowMsgBox "未正确指定卫材！": Me.txtMedi.SetFocus: Exit Sub
    gstrSQL = "": mstrTemp = ""
    With Me.msfMember
        For mintCount = 1 To .Rows - 1
            If .RowData(mintCount) <> 0 Then
                If Val(.TextMatrix(mintCount, col采用量)) = 0 Then
                    MsgBox mintCount & "行组成卫材的采用量没有输入！", vbInformation, gstrSysName: .SetFocus: Exit Sub
                End If
                If InStr(1, mstrTemp & ";", ";" & .RowData(mintCount) & ";") > 0 Then
                    MsgBox mintCount & "行卫材与前面发生重复！", vbInformation, gstrSysName: .SetFocus: Exit Sub
                End If
                mstrTemp = mstrTemp & ";" & .RowData(mintCount)
                gstrSQL = gstrSQL & "|" & .RowData(mintCount) & "^" & Val(.TextMatrix(mintCount, col采用量))
            End If
        Next
    End With
    
    If gstrSQL <> "" Then gstrSQL = Mid(gstrSQL, 2)
    
    gstrSQL = "zl_自制材料构成_UPDATE(" & Val(Me.lblMedi.Tag) & ",'" & gstrSQL & "')"
    err = 0: On Error GoTo ErrHand
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
                    
    ShowMsgBox Me.txtMedi.Text & Me.msfMember.Tag & "保存成功！"
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub CmdHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdMedi_Click()
    Dim rsTemp As New ADODB.Recordset
    
    err = 0: On Error GoTo ErrHand
    gstrSQL = "select I.ID,I.编码,I.名称,I.规格,I.产地,F.计算单位 as 单位" & _
            " from 收费项目目录 I,材料特性 S,诊疗项目目录 F" & _
            " where I.ID=S.材料ID and S.诊疗ID=F.ID and  I.类别='4'" & _
            "       and (I.撤档时间 is null or I.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))"
        
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
    
    With rsTemp
        If .BOF Or .EOF = 1 Then
            MsgBox "尚未建立该类具体规格的卫材！", vbExclamation, gstrSysName
            Me.lblMedi.Tag = 0: Me.txtMedi.Tag = "": Me.txtMedi.Text = Me.txtMedi.Tag: Me.txtMedi.SetFocus: Exit Sub
        End If
        If .RecordCount = 1 Then
            If Me.lblMedi.Tag <> !Id Then
                Me.lblMedi.Tag = !Id
                Me.txtMedi.Tag = "[" & !编码 & "]" & !名称
                Me.txtMedi.Text = Me.txtMedi.Tag
                If Me.Tag <> "7" Then
                    Me.lblSpec.Caption = "规格：" & IIf(IsNull(!规格), "", !规格) & _
                        "   厂牌：" & IIf(IsNull(!产地), "", !产地) & _
                        "   单位：" & IIf(IsNull(!单位), "", !单位)
                Else
                    Me.lblSpec.Caption = "产地：" & IIf(IsNull(!产地), "", !产地) & "   单位：" & IIf(IsNull(!单位), "", !单位)
                End If
                Call zlMemberRef(Me.lblMedi.Tag)
            End If
            Call OS.PressKey(vbKeyTab)
            Exit Sub
        End If
        Me.lvwItems.ListItems.Clear
        Do While Not .EOF
            Set mobjItem = Me.lvwItems.ListItems.Add(, "_" & !Id, !名称)
            mobjItem.Icon = "ItemUse": mobjItem.SmallIcon = "ItemUse"
            mobjItem.SubItems(Me.lvwItems.ColumnHeaders("编码").Index - 1) = !编码
            If Me.Tag <> "7" Then
                mobjItem.SubItems(Me.lvwItems.ColumnHeaders("规格").Index - 1) = IIf(IsNull(!规格), "", !规格)
            End If
            mobjItem.SubItems(Me.lvwItems.ColumnHeaders("产地").Index - 1) = IIf(IsNull(!产地), "", !产地)
            mobjItem.SubItems(Me.lvwItems.ColumnHeaders("单位").Index - 1) = IIf(IsNull(!单位), "", !单位)
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

    Dim rsTemp As New ADODB.Recordset
    Me.Caption = "自制卫材构成"
    Me.lblNote.Caption = "    选择具体规格的卫材，按剂量单位指定其具体的原料卫材；" & _
            "未指定其原料卫材（或清除其所有原料），本卫材将不会作为自制卫材。"
    Me.lblMember.Caption = "原料卫材(&E)："
    
    If Me.cmdClose.Tag = "查阅" Then
        Me.msfMember.Active = False
        Me.CmdSave.Visible = False
        Me.cmdClear.Visible = False
        Me.cmdRestore.Visible = False
    Else
        Me.msfMember.Active = True
    End If
    
    err = 0: On Error GoTo ErrHand
    
    gstrSQL = "select I.ID,I.编码,I.名称,I.规格,I.产地,F.计算单位 as 单位" & _
            " from 收费项目目录 I,材料特性 S,诊疗项目目录 F" & _
            " where I.ID=S.材料ID and S.诊疗ID=F.ID and I.类别='4'and I.ID=[1]" & _
            "       and (I.撤档时间 is null or I.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))"
            
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(Me.lblMedi.Tag))
            
    With rsTemp
        If .BOF Or .EOF = 1 Then
            Me.lblMedi.Tag = 0: Me.txtMedi.Tag = "": Me.txtMedi.Text = Me.txtMedi.Tag
        Else
            Me.lblMedi.Tag = !Id
            Me.txtMedi.Tag = "[" & !编码 & "]" & !名称
            Me.txtMedi.Text = Me.txtMedi.Tag
            Me.lblSpec.Caption = "规格：" & IIf(IsNull(!规格), "", !规格) & _
                "   厂牌：" & IIf(IsNull(!产地), "", !产地) & _
                "   单位：" & IIf(IsNull(!单位), "", !单位)
            Call zlMemberRef(Me.lblMedi.Tag)
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
            Me.msfMember.SetFocus
        End If
    Else
        cmdClose_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    With Me.msfMember
        .MsfObj.FixedCols = 1: .Cols = 6

        .TextMatrix(0, 0) = "": .TextMatrix(0, col品名) = "卫材名称"
        .TextMatrix(0, col规格) = "规格": .TextMatrix(0, col产地) = "产地"
        .TextMatrix(0, col采用量) = "采用量": .TextMatrix(0, col单位) = "单位"
        
        .ColAlignment(col品名) = 1: .ColAlignment(col规格) = 1: .ColAlignment(col产地) = 1: .ColAlignment(col单位) = 7
        
        .ColWidth(0) = 300: .ColWidth(col品名) = 2800
        .ColWidth(col规格) = 1200: .ColWidth(col产地) = 1200: .ColWidth(col采用量) = 1000: .ColWidth(col单位) = 800

        .ColData(0) = 5: .ColData(col品名) = 1
        .ColData(col规格) = 5: .ColData(col产地) = 5: .ColData(col采用量) = 4: .ColData(col单位) = 5
        
        .PrimaryCol = col品名: .LocateCol = col品名
        .TextMatrix(1, 0) = "1": .Row = 1: .Col = col品名
    End With
    
    Me.lvwItems.ListItems.Clear
    With Me.lvwItems.ColumnHeaders
        .Clear
        .Add , "名称", "名称", 2000
        .Add , "编码", "编码", 1000
        .Add , "规格", "规格", 1200
        .Add , "产地", "产地", 1200
        .Add , "单位", "单位", 600
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
                If Me.Tag <> "7" Then
                    Me.lblSpec.Caption = "规格：" & .SelectedItem.SubItems(.ColumnHeaders("规格").Index - 1) & _
                        "   厂牌：" & .SelectedItem.SubItems(.ColumnHeaders("产地").Index - 1) & _
                        "   单位：" & .SelectedItem.SubItems(.ColumnHeaders("单位").Index - 1)
                Else
                    Me.lblSpec.Caption = "产地：" & .SelectedItem.SubItems(.ColumnHeaders("产地").Index - 1) & _
                        "   单位：" & .SelectedItem.SubItems(.ColumnHeaders("单位").Index - 1)
                End If
                Call zlMemberRef(Me.lblMedi.Tag)
            End If
            Me.txtMedi.SetFocus
            Call OS.PressKey(vbKeyTab)
        Else
            Me.msfMember.RowData(Me.msfMember.Row) = Mid(.SelectedItem.Key, 2)
            Me.msfMember.Text = "[" & .SelectedItem.SubItems(.ColumnHeaders("编码").Index - 1) & "]" & .SelectedItem.Text
            Me.msfMember.TextMatrix(Me.msfMember.Row, col品名) = Me.msfMember.Text
            Me.msfMember.TextMatrix(Me.msfMember.Row, col规格) = .SelectedItem.SubItems(.ColumnHeaders("规格").Index - 1)
            Me.msfMember.TextMatrix(Me.msfMember.Row, col产地) = .SelectedItem.SubItems(.ColumnHeaders("产地").Index - 1)
            Me.msfMember.TextMatrix(Me.msfMember.Row, col单位) = .SelectedItem.SubItems(.ColumnHeaders("单位").Index - 1)
            Me.msfMember.SetFocus
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

Private Sub msfMember_AfterAddRow(Row As Long)
    With Me.msfMember
        For mintCount = Row To .Rows - 1
            .TextMatrix(mintCount, 0) = mintCount
        Next
    End With
End Sub

Private Sub msfMember_AfterDeleteRow()
    With Me.msfMember
        For mintCount = IIf(.Row <> 1, .Row - 1, .Row) To .Rows - 1
            .TextMatrix(mintCount, 0) = mintCount
        Next
    End With
End Sub

Private Sub msfMember_CommandClick()
    Dim rsTemp As New ADODB.Recordset
    
    err = 0: On Error GoTo ErrHand
        gstrSQL = "select I.ID,I.编码,I.名称,I.规格,I.产地,I.计算单位 as 单位" & _
                " from 收费项目目录 I,材料特性 S" & _
                " where I.ID=S.材料ID   " & _
                "       and (I.撤档时间 is null or I.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))" & _
                "       and I.ID<>[1] "
        gstrSQL = gstrSQL & "      and I.类别='4' and S.原材料=1"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(Me.lblMedi.Tag))
       
    With rsTemp
        If .BOF Or .EOF = 1 Then
            MsgBox "尚未建立可作为组成卫生材料的规格！", vbExclamation, gstrSysName
            Me.lblMedi.Tag = 0: Me.txtMedi.Tag = "": Me.txtMedi.Text = Me.txtMedi.Tag: Me.txtMedi.SetFocus: Exit Sub
        End If
        
        If .RecordCount = 1 Then
            Me.msfMember.RowData(Me.msfMember.Row) = !Id
            Me.msfMember.Text = "[" & !编码 & "]" & !名称
            Me.msfMember.TextMatrix(Me.msfMember.Row, col品名) = Me.msfMember.Text
            Me.msfMember.TextMatrix(Me.msfMember.Row, col规格) = IIf(IsNull(!规格), "", !规格)
            Me.msfMember.TextMatrix(Me.msfMember.Row, col产地) = IIf(IsNull(!产地), "", !产地)
            Me.msfMember.TextMatrix(Me.msfMember.Row, col单位) = IIf(IsNull(!单位), "", !单位)
            Exit Sub
        End If
        Me.lvwItems.ListItems.Clear
        Do While Not .EOF
            Set mobjItem = Me.lvwItems.ListItems.Add(, "_" & !Id, !名称)
            mobjItem.Icon = "ItemUse": mobjItem.SmallIcon = "ItemUse"
            mobjItem.SubItems(Me.lvwItems.ColumnHeaders("编码").Index - 1) = !编码
            mobjItem.SubItems(Me.lvwItems.ColumnHeaders("规格").Index - 1) = IIf(IsNull(!规格), "", !规格)
            mobjItem.SubItems(Me.lvwItems.ColumnHeaders("产地").Index - 1) = IIf(IsNull(!产地), "", !产地)
            mobjItem.SubItems(Me.lvwItems.ColumnHeaders("单位").Index - 1) = IIf(IsNull(!单位), "", !单位)
            .MoveNext
        Loop
        Me.lvwItems.ListItems(1).Selected = True
    End With
    With Me.lvwItems
        .Tag = Me.msfMember.Name
        .Left = Me.msfMember.Left + 500
        .Top = Me.msfMember.Top
        .ZOrder 0: .Visible = True
        .SetFocus
    End With
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub msfMember_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim rsTemp As New ADODB.Recordset
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    With Me.msfMember
        If .TxtVisible = False Then Exit Sub
        If .Col <> 1 Then
            If Trim(.Text) = "" Then
                MsgBox "请输入采用量！", vbInformation, gstrSysName: Cancel = True: .TxtSetFocus: Exit Sub
            End If
            If Not IsNumeric(.Text) Then
                MsgBox "采用量中含有非法字符！", vbInformation, gstrSysName: Cancel = True: .TxtSetFocus: Exit Sub
            End If
            If Val(.Text) > 10000000 Then
                MsgBox "采用量超过最大值！", vbInformation, gstrSysName: Cancel = True: .TxtSetFocus: Exit Sub
            End If
            If Val(.Text) <= 0 Then
                MsgBox "采用量必须大于0！", vbInformation, gstrSysName: Cancel = True: .TxtSetFocus: Exit Sub
            End If
            .Text = Format(.Text, "0.000"): .TextMatrix(.Row, col采用量) = .Text
            Exit Sub
        End If
    End With
    
    mstrTemp = UCase(Trim(Me.msfMember.Text))
    If InStr(1, mstrTemp, "[") <> 0 And InStr(1, mstrTemp, "]") <> 0 Then mstrTemp = Mid(mstrTemp, 2, InStr(1, mstrTemp, "]") - 2)
    
    err = 0: On Error GoTo ErrHand
    mstrTemp = GetMatchingSting(mstrTemp)
    gstrSQL = "select distinct I.ID,I.编码,I.名称,I.规格,I.产地,F.计算单位 as 单位" & _
            " from 收费项目目录 I,收费项目别名 N,材料特性 S,诊疗项目目录 F" & _
            " where I.ID=S.材料ID and S.诊疗ID=F.ID and I.ID=N.收费细目ID " & _
            "       and (I.撤档时间 is null or I.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))" & _
            "       and (I.编码 like [2] or N.名称 like [2] or N.简码 like [2])" & _
            "       and I.ID<>[1] "
                    
    gstrSQL = gstrSQL & "      and I.类别='4' and S.原材料=1"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(Me.lblMedi.Tag), mstrTemp)
    
      
    With rsTemp
        If .EOF Then
            MsgBox "未找到相关卫材，请重新输入！", vbInformation, gstrSysName: Cancel = True: Me.msfMember.TxtSetFocus
            Exit Sub
        End If
        If .RecordCount = 1 Then
            Me.msfMember.RowData(Me.msfMember.Row) = !Id
            Me.msfMember.Text = "[" & !编码 & "]" & !名称
            Me.msfMember.TextMatrix(Me.msfMember.Row, col品名) = Me.msfMember.Text
            Me.msfMember.TextMatrix(Me.msfMember.Row, col规格) = IIf(IsNull(!规格), "", !规格)
            Me.msfMember.TextMatrix(Me.msfMember.Row, col产地) = IIf(IsNull(!产地), "", !产地)
            Me.msfMember.TextMatrix(Me.msfMember.Row, col单位) = IIf(IsNull(!单位), "", !单位)
            Exit Sub
        End If
        Me.lvwItems.ListItems.Clear
        Do While Not .EOF
            Set mobjItem = Me.lvwItems.ListItems.Add(, "_" & !Id, !名称)
            mobjItem.Icon = "ItemUse": mobjItem.SmallIcon = "ItemUse"
            mobjItem.SubItems(Me.lvwItems.ColumnHeaders("编码").Index - 1) = !编码
            mobjItem.SubItems(Me.lvwItems.ColumnHeaders("规格").Index - 1) = IIf(IsNull(!规格), "", !规格)
            mobjItem.SubItems(Me.lvwItems.ColumnHeaders("产地").Index - 1) = IIf(IsNull(!产地), "", !产地)
            mobjItem.SubItems(Me.lvwItems.ColumnHeaders("单位").Index - 1) = IIf(IsNull(!单位), "", !单位)
            .MoveNext
        Loop
        Me.lvwItems.ListItems(1).Selected = True
    End With
    With Me.lvwItems
        .Tag = Me.msfMember.Name
        .Left = Me.msfMember.Left + 500
        .Top = Me.msfMember.Top
        .ZOrder 0: .Visible = True
        .SetFocus
    End With
    Cancel = True: Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtMedi_GotFocus()
    Me.txtMedi.SelStart = 0: Me.txtMedi.SelLength = 100
End Sub

Private Sub txtMedi_KeyPress(KeyAscii As Integer)
    Dim rsTemp As New ADODB.Recordset
    
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If KeyAscii <> vbKeyReturn Then Exit Sub
    mstrTemp = UCase(Trim(Me.txtMedi.Text))
    If mstrTemp = "" Then Me.lblMedi.Tag = 0: Me.txtMedi.Tag = "": Me.txtMedi.Text = "": Exit Sub
    
    If InStr(1, mstrTemp, "[") <> 0 And InStr(1, mstrTemp, "]") <> 0 Then mstrTemp = Mid(mstrTemp, 2, InStr(1, mstrTemp, "]") - 2)
    err = 0: On Error GoTo ErrHand
    mstrTemp = GetMatchingSting(mstrTemp)
    
    gstrSQL = "select distinct I.ID,I.编码,I.名称,I.规格,I.产地,I.计算单位 as 单位" & _
            " from 收费项目目录 I,收费项目别名 N,材料特性 S,诊疗项目目录 F" & _
            " where I.ID=S.材料ID and S.诊疗ID=F.ID and I.ID=N.收费细目ID and I.类别='4'" & _
            "       and (I.撤档时间 is null or I.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))" & _
            "       and (I.编码 like [1] or N.名称 like [1] or N.简码 like [1])"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mstrTemp)
       
    With rsTemp
        If .BOF Or .EOF = 1 Then
            MsgBox "未找到指定规格的材料，请重新指定！", vbExclamation, gstrSysName
            Me.lblMedi.Tag = 0: Me.txtMedi.Tag = "": Me.txtMedi.Text = Me.txtMedi.Tag: Me.txtMedi.SetFocus: Exit Sub
        End If
        
        If .RecordCount = 1 Then
            If Me.lblMedi.Tag <> !Id Then
                Me.lblMedi.Tag = !Id
                Me.txtMedi.Tag = "[" & !编码 & "]" & !名称
                Me.txtMedi.Text = Me.txtMedi.Tag
                Me.lblSpec.Caption = "规格：" & IIf(IsNull(!规格), "", !规格) & _
                    "   产地：" & IIf(IsNull(!产地), "", !产地) & _
                    "   单位：" & IIf(IsNull(!单位), "", !单位)
                Call zlMemberRef(Me.lblMedi.Tag)
            End If
            Call OS.PressKey(vbKeyTab)
            Exit Sub
        End If
        Me.lvwItems.ListItems.Clear
        Do While Not .EOF
            Set mobjItem = Me.lvwItems.ListItems.Add(, "_" & !Id, !名称)
            mobjItem.Icon = "ItemUse": mobjItem.SmallIcon = "ItemUse"
            mobjItem.SubItems(Me.lvwItems.ColumnHeaders("编码").Index - 1) = !编码
            If Me.Tag <> "7" Then
                mobjItem.SubItems(Me.lvwItems.ColumnHeaders("规格").Index - 1) = IIf(IsNull(!规格), "", !规格)
            End If
            mobjItem.SubItems(Me.lvwItems.ColumnHeaders("产地").Index - 1) = IIf(IsNull(!产地), "", !产地)
            mobjItem.SubItems(Me.lvwItems.ColumnHeaders("单位").Index - 1) = IIf(IsNull(!单位), "", !单位)
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

Private Sub zlMemberRef(lngMediId As Long)
    '--------------------------------------------------------
    '功能：刷新指定卫材的协定组成卫材或原料卫材
    '入参：lngMediId-指定的药名id
    '--------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    
    err = 0: On Error GoTo ErrHand
    gstrSQL = "select I.ID,I.编码,I.名称,I.规格,I.产地,M.计算单位 as 单位,P.分子 as 采用量" & _
        " from 自制材料构成 P,收费项目目录 I,材料特性 S,诊疗项目目录 M" & _
        " where P.原料材料ID=I.ID and I.ID=S.材料ID and S.诊疗id=M.ID" & _
        "       and P.自制材料ID=[1]"

    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngMediId)
    
     
    With rsTemp
        Me.msfMember.ClearBill
        Do While Not .EOF
            If Me.msfMember.Rows < .AbsolutePosition + 1 Then Me.msfMember.Rows = Me.msfMember.Rows + 1
            Me.msfMember.RowData(.AbsolutePosition) = !Id
            Me.msfMember.TextMatrix(.AbsolutePosition, 0) = .AbsolutePosition
            Me.msfMember.TextMatrix(.AbsolutePosition, col品名) = "[" & !编码 & "]" & !名称
            Me.msfMember.TextMatrix(.AbsolutePosition, col规格) = IIf(IsNull(!规格), "", !规格)
            Me.msfMember.TextMatrix(.AbsolutePosition, col产地) = IIf(IsNull(!产地), "", !产地)
            Me.msfMember.TextMatrix(.AbsolutePosition, col采用量) = Format(!采用量, "0.000")
            Me.msfMember.TextMatrix(.AbsolutePosition, col单位) = IIf(IsNull(!单位), "", !单位)
            .MoveNext
        Loop
    End With
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

