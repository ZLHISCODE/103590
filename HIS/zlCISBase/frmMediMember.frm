VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmMediMember 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "药品组成"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7875
   Icon            =   "frmMediMember.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   7875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdMedi 
      Caption         =   "…"
      Height          =   285
      Left            =   7440
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   788
      Width           =   285
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "关闭(&X)"
      Height          =   350
      Left            =   6625
      TabIndex        =   5
      Top             =   3780
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   90
      Picture         =   "frmMediMember.frx":058A
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3780
      Width           =   1100
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "保存(&S)"
      Height          =   350
      Left            =   5535
      TabIndex        =   4
      Top             =   3780
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
      Picture         =   "frmMediMember.frx":06D4
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3780
      Width           =   1290
   End
   Begin VB.CommandButton cmdRestore 
      Caption         =   "恢复(&R)"
      Height          =   350
      Left            =   2565
      Picture         =   "frmMediMember.frx":081E
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3780
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
            Picture         =   "frmMediMember.frx":0968
            Key             =   "ItemUse"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediMember.frx":0F02
            Key             =   "ItemStop"
         EndProperty
      EndProperty
   End
   Begin ZL9BillEdit.BillEdit msfMember 
      Height          =   2055
      Left            =   90
      TabIndex        =   3
      Top             =   1620
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   3625
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
      Top             =   4425
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
      Caption         =   "规格：      生产商：       单位：瓶"
      Height          =   180
      Left            =   1125
      TabIndex        =   11
      Top             =   1110
      Width           =   3150
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   75
      Picture         =   "frmMediMember.frx":149C
      Top             =   90
      Width           =   480
   End
   Begin VB.Label lblnote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "    选择具体规格的药品，按剂量单位指定其具体的组成药品；未指定其组成（或清除其组成），将不会作为协定药品。"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   720
      TabIndex        =   10
      Top             =   150
      Width           =   7065
   End
   Begin VB.Label lblMedi 
      AutoSize        =   -1  'True
      Caption         =   "指定药品(&M)"
      Height          =   180
      Left            =   90
      TabIndex        =   0
      Top             =   840
      Width           =   990
   End
   Begin VB.Label lblMember 
      AutoSize        =   -1  'True
      Caption         =   "组成药品(&E)："
      Height          =   180
      Left            =   90
      TabIndex        =   2
      Top             =   1395
      Width           =   1170
   End
End
Attribute VB_Name = "frmMediMember"
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
'   4、当前编辑内容：由Me.msfMember.Tag保存，分别为"协定"、"自制"
'---------------------------------------------------
'协定组成药品只能为同材质的药品；
'自制原料药品的材质关系:
'   西成药：自制原料药品只能为“西成药”的原料类药物
'   中成药：自制原料药品可以为“中成药”和“中草药”的原料药物
'   中草药：自制原料药品只能为“中草药”的原料药物
'---------------------------------------------------
Dim rsTemp As New ADODB.Recordset
Dim objItem As ListItem
Dim strTemp As String
Dim intCount As Integer

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

Private Sub cmdSave_Click()
    If Val(Me.lblMedi.Tag) = 0 Then MsgBox "未正确指定药品！", vbExclamation, gstrSysName: Me.txtMedi.SetFocus: Exit Sub
    gstrSql = "": strTemp = ""
    With Me.msfMember
        For intCount = 1 To .Rows - 1
            If .RowData(intCount) <> 0 Then
                If Val(.TextMatrix(intCount, col采用量)) = 0 Then
                    MsgBox intCount & "行组成药品的采用量没有输入！", vbInformation, gstrSysName: .SetFocus: Exit Sub
                End If
                If InStr(1, strTemp & ";", ";" & .RowData(intCount) & ";") > 0 Then
                    MsgBox intCount & "行药品与前面发生重复！", vbInformation, gstrSysName: .SetFocus: Exit Sub
                End If
                strTemp = strTemp & ";" & .RowData(intCount)
                gstrSql = gstrSql & "|" & .RowData(intCount) & "^" & Val(.TextMatrix(intCount, col采用量))
            End If
        Next
    End With
    If gstrSql <> "" Then gstrSql = Mid(gstrSql, 2)
    If Me.msfMember.Tag = "协定" Then
        gstrSql = "zl_协定药品对照_UPDATE(" & Val(Me.lblMedi.Tag) & ",'" & gstrSql & "')"
    Else
        gstrSql = "zl_自制药品构成_UPDATE(" & Val(Me.lblMedi.Tag) & ",'" & gstrSql & "')"
    End If
    Err = 0: On Error GoTo ErrHand
    Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
    MsgBox Me.txtMedi.Text & Me.msfMember.Tag & "保存成功！", vbExclamation, gstrSysName
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
    
    gstrSql = "select I.ID,I.编码,I.名称,I.规格,I.产地,F.计算单位 as 单位" & _
            " from 收费项目目录 I,药品规格 S,诊疗项目目录 F" & _
            " where I.ID=S.药品ID and S.药名ID=F.ID and  I.类别=[1] " & _
            "       and (I.撤档时间 is null or I.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Me.Tag)
    
    With rsTemp
        If .BOF Or .EOF Then
            MsgBox "尚未建立该类具体规格的药品！", vbExclamation, gstrSysName
            Me.lblMedi.Tag = 0: Me.txtMedi.Tag = "": Me.txtMedi.Text = Me.txtMedi.Tag: Me.txtMedi.SetFocus: Exit Sub
        End If
        If .RecordCount = 1 Then
            If Me.lblMedi.Tag <> !ID Then
                Me.lblMedi.Tag = !ID
                Me.txtMedi.Tag = "[" & !编码 & "]" & !名称
                Me.txtMedi.Text = Me.txtMedi.Tag
                If Me.Tag <> "7" Then
                    Me.lblSpec.Caption = "规格：" & IIf(IsNull(!规格), "", !规格) & _
                        "   生产商：" & IIf(IsNull(!产地), "", !产地) & _
                        "   剂量单位：" & IIf(IsNull(!单位), "", !单位)
                Else
                    Me.lblSpec.Caption = "规格：" & IIf(IsNull(!规格), "", !规格) & _
                        "   生产商：" & IIf(IsNull(!产地), "", !产地) & _
                        "   剂量单位：" & IIf(IsNull(!单位), "", !单位)
                End If
                Call zlMemberRef(Me.lblMedi.Tag)
            End If
            Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        Me.lvwItems.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !名称)
            objItem.Icon = "ItemUse": objItem.SmallIcon = "ItemUse"
            objItem.SubItems(Me.lvwItems.ColumnHeaders("编码").Index - 1) = !编码
'            If Me.Tag <> "7" Then
                objItem.SubItems(Me.lvwItems.ColumnHeaders("规格").Index - 1) = IIf(IsNull(!规格), "", !规格)
'            End If
            objItem.SubItems(Me.lvwItems.ColumnHeaders("产地").Index - 1) = IIf(IsNull(!产地), "", !产地)
            objItem.SubItems(Me.lvwItems.ColumnHeaders("单位").Index - 1) = IIf(IsNull(!单位), "", !单位)
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
    With Me.msfMember
        .MsfObj.FixedCols = 1: .Cols = 6

        .TextMatrix(0, 0) = "": .TextMatrix(0, col品名) = "药品名称"
        .TextMatrix(0, col规格) = "规格"
        If Me.Tag <> "7" Then
            .TextMatrix(0, col产地) = "生产商"
        Else
            .TextMatrix(0, col产地) = "生产商"
        End If
        .TextMatrix(0, col采用量) = "采用量": .TextMatrix(0, col单位) = "剂量单位"
        
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
        If Me.Tag <> "7" Then
            .Add , "产地", "生产商", 1200
        Else
            .Add , "产地", "生产商", 1200
        End If
        .Add , "单位", "剂量单位", 600
    End With
    With Me.lvwItems
        .ColumnHeaders("编码").Position = 1
        .SortKey = .ColumnHeaders("编码").Index - 1
        .SortOrder = lvwAscending
    End With
    
    
    If Me.msfMember.Tag = "协定" Then
        Me.Caption = "协定药品组成"
        Me.lblnote.Caption = "    选择具体规格的药品，按剂量单位指定其具体的组成药品；" & _
                "未指定其组成（或清除其组成），本药品将不会作为协定药品。"
        Me.lblMember.Caption = "组成药品(&E)："
    Else
        Me.Caption = "自制药品构成"
        Me.lblnote.Caption = "    选择具体规格的药品，按剂量单位指定其具体的原料药品；" & _
                "未指定其原料药品（或清除其所有原料），本药品将不会作为自制药品。"
        Me.lblMember.Caption = "原料药品(&E)："
    End If
    If Me.cmdClose.Tag = "查阅" Then
        Me.msfMember.Active = False
        Me.cmdSave.Visible = False
        Me.cmdClear.Visible = False
        Me.cmdRestore.Visible = False
    Else
        Me.msfMember.Active = True
    End If
    
    Err = 0: On Error GoTo ErrHand
    
    gstrSql = "select I.ID,I.编码,I.名称,I.规格,I.产地,F.计算单位 as 单位" & _
            " from 收费项目目录 I,药品规格 S,诊疗项目目录 F" & _
            " where I.ID=S.药品ID and S.药名ID=F.ID and I.类别=[1] and I.ID=[2] " & _
            "       and (I.撤档时间 is null or I.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Me.Tag, Val(Me.lblMedi.Tag))
    
    With rsTemp
        If .BOF Or .EOF Then
            Me.lblMedi.Tag = 0: Me.txtMedi.Tag = "": Me.txtMedi.Text = Me.txtMedi.Tag
        Else
            Me.lblMedi.Tag = !ID
            Me.txtMedi.Tag = "[" & !编码 & "]" & !名称
            Me.txtMedi.Text = Me.txtMedi.Tag
            If Me.Tag <> "7" Then
                Me.lblSpec.Caption = "规格：" & IIf(IsNull(!规格), "", !规格) & _
                    "   生产商：" & IIf(IsNull(!产地), "", !产地) & _
                    "   剂量单位：" & IIf(IsNull(!单位), "", !单位)
            Else
                Me.lblSpec.Caption = "规格：" & IIf(IsNull(!规格), "", !规格) & _
                    "   生产商：" & IIf(IsNull(!产地), "", !产地) & _
                    "   剂量单位：" & IIf(IsNull(!单位), "", !单位)
            End If
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
'
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
                        "   生产商：" & .SelectedItem.SubItems(.ColumnHeaders("产地").Index - 1) & _
                        "   剂量单位：" & .SelectedItem.SubItems(.ColumnHeaders("单位").Index - 1)
                Else
                    Me.lblSpec.Caption = "规格：" & .SelectedItem.SubItems(.ColumnHeaders("规格").Index - 1) & _
                        "   生产商：" & .SelectedItem.SubItems(.ColumnHeaders("产地").Index - 1) & _
                        "   剂量单位：" & .SelectedItem.SubItems(.ColumnHeaders("单位").Index - 1)
                End If
                Call zlMemberRef(Me.lblMedi.Tag)
            End If
            Me.txtMedi.SetFocus
            Call zlCommFun.PressKey(vbKeyTab)
        Else
            Me.msfMember.RowData(Me.msfMember.Row) = Mid(.SelectedItem.Key, 2)
            Me.msfMember.Text = "[" & .SelectedItem.SubItems(.ColumnHeaders("编码").Index - 1) & "]" & .SelectedItem.Text
            Me.msfMember.TextMatrix(Me.msfMember.Row, 0) = msfMember.Rows - 1
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
        For intCount = Row To .Rows - 1
            .TextMatrix(intCount, 0) = intCount
        Next
    End With
End Sub

Private Sub msfMember_AfterDeleteRow()
    With Me.msfMember
        For intCount = IIf(.Row <> 1, .Row - 1, .Row) To .Rows - 1
            .TextMatrix(intCount, 0) = intCount
        Next
    End With
End Sub

Private Sub msfMember_CommandClick()
    Err = 0: On Error GoTo ErrHand
    
    If Me.msfMember.Tag = "协定" Then
        gstrSql = "select I.ID,I.编码,I.名称,I.规格,I.产地,F.计算单位 as 单位" & _
                " from 收费项目目录 I,药品规格 S,诊疗项目目录 F" & _
                " where I.ID=S.药品ID and S.药名ID=F.ID and  I.类别=[1] " & _
                "       and (I.撤档时间 is null or I.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))" & _
                "       and I.ID<>[2] "
    Else
        gstrSql = "select I.ID,I.编码,I.名称,I.规格,I.产地,F.计算单位 as 单位" & _
                " from 收费项目目录 I,药品规格 S,诊疗项目目录 F,药品特性 T" & _
                " where I.ID=S.药品ID and S.药名ID=F.ID and F.ID=T.药名ID" & _
                "       and (I.撤档时间 is null or I.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))" & _
                "       and I.ID<>[2] "
        Select Case Me.Tag
        Case "5"
            gstrSql = gstrSql & "      and I.类别='5' and T.是否原料=1"
        Case "6"
            gstrSql = gstrSql & "      and I.类别 in ('6','7') and T.是否原料=1"
        Case "7"
            gstrSql = gstrSql & "      and I.类别='7' and T.是否原料=1"
        End Select
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Me.Tag, Val(Me.lblMedi.Tag))
    
    With rsTemp
        If .BOF Or .EOF Then
            MsgBox "尚未建立可作为组成药物的药品规格！", vbExclamation, gstrSysName
            Me.lblMedi.Tag = 0: Me.txtMedi.Tag = "": Me.txtMedi.Text = Me.txtMedi.Tag: Me.txtMedi.SetFocus: Exit Sub
        End If
        If .RecordCount = 1 Then
            Me.msfMember.RowData(Me.msfMember.Row) = !ID
            Me.msfMember.Text = "[" & !编码 & "]" & !名称
            Me.msfMember.TextMatrix(Me.msfMember.Row, col品名) = Me.msfMember.Text
            Me.msfMember.TextMatrix(Me.msfMember.Row, col规格) = IIf(IsNull(!规格), "", !规格)
            Me.msfMember.TextMatrix(Me.msfMember.Row, col产地) = IIf(IsNull(!产地), "", !产地)
            Me.msfMember.TextMatrix(Me.msfMember.Row, col单位) = IIf(IsNull(!单位), "", !单位)
            Exit Sub
        End If
        Me.lvwItems.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !名称)
            objItem.Icon = "ItemUse": objItem.SmallIcon = "ItemUse"
            objItem.SubItems(Me.lvwItems.ColumnHeaders("编码").Index - 1) = !编码
            objItem.SubItems(Me.lvwItems.ColumnHeaders("规格").Index - 1) = IIf(IsNull(!规格), "", !规格)
            objItem.SubItems(Me.lvwItems.ColumnHeaders("产地").Index - 1) = IIf(IsNull(!产地), "", !产地)
            objItem.SubItems(Me.lvwItems.ColumnHeaders("单位").Index - 1) = IIf(IsNull(!单位), "", !单位)
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
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub msfMember_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
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
            .Text = Format(.Text, "0.00000"): .TextMatrix(.Row, col采用量) = .Text
            Exit Sub
        End If
    End With
    
    strTemp = UCase(Trim(Me.msfMember.Text))
    If InStr(1, strTemp, "[") <> 0 And InStr(1, strTemp, "]") <> 0 Then strTemp = Mid(strTemp, 2, InStr(1, strTemp, "]") - 2)
    
    Err = 0: On Error GoTo ErrHand
    
    If Me.msfMember.Tag = "协定" Then
        gstrSql = "select distinct I.ID,I.编码,I.名称,I.规格,I.产地,F.计算单位 as 单位" & _
                " from 收费项目目录 I,收费项目别名 N,药品规格 S,诊疗项目目录 F" & _
                " where I.ID=S.药品ID and S.药名ID=F.ID and I.ID=N.收费细目ID and I.类别=[1] " & _
                "       and (I.撤档时间 is null or I.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))" & _
                "       and (I.编码 like [2] or N.名称 like [3] or N.简码 like [3])" & _
                "       and I.ID<>[4] "
    Else
        gstrSql = "select distinct I.ID,I.编码,I.名称,I.规格,I.产地,F.计算单位 as 单位" & _
                " from 收费项目目录 I,收费项目别名 N,药品规格 S,诊疗项目目录 F,药品特性 T" & _
                " where I.ID=S.药品ID and S.药名ID=F.ID and I.ID=N.收费细目ID and F.ID=T.药名ID" & _
                "       and (I.撤档时间 is null or I.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))" & _
                "       and (I.编码 like [2] or N.名称 like [3] or N.简码 like [3])" & _
                "       and I.ID<>[4] "
        Select Case Me.Tag
        Case "5"
            gstrSql = gstrSql & "      and I.类别='5' and T.是否原料=1"
        Case "6"
            gstrSql = gstrSql & "      and I.类别 in ('6','7') and T.是否原料=1"
        Case "7"
            gstrSql = gstrSql & "      and I.类别='7' and T.是否原料=1"
        End Select
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Me.Tag, strTemp & "%", gstrMatch & strTemp & "%", Val(Me.lblMedi.Tag))
    
    With rsTemp
        If .EOF Then
            MsgBox "未找到相关药品，请重新输入！", vbInformation, gstrSysName: Cancel = True: Me.msfMember.TxtSetFocus
            Exit Sub
        End If
        If .RecordCount = 1 Then
            Me.msfMember.RowData(Me.msfMember.Row) = !ID
            Me.msfMember.Text = "[" & !编码 & "]" & !名称
            Me.msfMember.TextMatrix(Me.msfMember.Row, col品名) = Me.msfMember.Text
            Me.msfMember.TextMatrix(Me.msfMember.Row, col规格) = IIf(IsNull(!规格), "", !规格)
            Me.msfMember.TextMatrix(Me.msfMember.Row, col产地) = IIf(IsNull(!产地), "", !产地)
            Me.msfMember.TextMatrix(Me.msfMember.Row, col单位) = IIf(IsNull(!单位), "", !单位)
            Exit Sub
        End If
        Me.lvwItems.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !名称)
            objItem.Icon = "ItemUse": objItem.SmallIcon = "ItemUse"
            objItem.SubItems(Me.lvwItems.ColumnHeaders("编码").Index - 1) = !编码
            objItem.SubItems(Me.lvwItems.ColumnHeaders("规格").Index - 1) = IIf(IsNull(!规格), "", !规格)
            objItem.SubItems(Me.lvwItems.ColumnHeaders("产地").Index - 1) = IIf(IsNull(!产地), "", !产地)
            objItem.SubItems(Me.lvwItems.ColumnHeaders("单位").Index - 1) = IIf(IsNull(!单位), "", !单位)
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
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If KeyAscii <> vbKeyReturn Then Exit Sub
    strTemp = UCase(Trim(Me.txtMedi.Text))
    If strTemp = "" Then Me.lblMedi.Tag = 0: Me.txtMedi.Tag = "": Me.txtMedi.Text = "": Exit Sub
    
    If InStr(1, strTemp, "[") <> 0 And InStr(1, strTemp, "]") <> 0 Then strTemp = Mid(strTemp, 2, InStr(1, strTemp, "]") - 2)
    Err = 0: On Error GoTo ErrHand
    
    gstrSql = "select distinct I.ID,I.编码,I.名称,I.规格,I.产地,I.计算单位 as 单位" & _
            " from 收费项目目录 I,收费项目别名 N,药品规格 S,诊疗项目目录 F" & _
            " where I.ID=S.药品ID and S.药名ID=F.ID and I.ID=N.收费细目ID and I.类别=[1] " & _
            "       and (I.撤档时间 is null or I.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))" & _
            "       and (I.编码 like [2] or N.名称 like [3] or N.简码 like [3])"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Me.Tag, strTemp & "%", gstrMatch & strTemp & "%")
    
    With rsTemp
        If .BOF Or .EOF Then
            MsgBox "未找到指定规格的药品，请重新指定！", vbExclamation, gstrSysName
            Me.lblMedi.Tag = 0: Me.txtMedi.Tag = "": Me.txtMedi.Text = Me.txtMedi.Tag: Me.txtMedi.SetFocus: Exit Sub
        End If
        If .RecordCount = 1 Then
            If Me.lblMedi.Tag <> !ID Then
                Me.lblMedi.Tag = !ID
                Me.txtMedi.Tag = "[" & !编码 & "]" & !名称
                Me.txtMedi.Text = Me.txtMedi.Tag
                If Me.Tag <> "7" Then
                    Me.lblSpec.Caption = "规格：" & IIf(IsNull(!规格), "", !规格) & _
                        "   生产商：" & IIf(IsNull(!产地), "", !产地) & _
                        "   剂量单位：" & IIf(IsNull(!单位), "", !单位)
                Else
                    Me.lblSpec.Caption = "规格：" & IIf(IsNull(!规格), "", !规格) & _
                        "   生产商：" & IIf(IsNull(!产地), "", !产地) & _
                        "   剂量单位：" & IIf(IsNull(!单位), "", !单位)
                End If
                Call zlMemberRef(Me.lblMedi.Tag)
            End If
            Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        Me.lvwItems.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !名称)
            objItem.Icon = "ItemUse": objItem.SmallIcon = "ItemUse"
            objItem.SubItems(Me.lvwItems.ColumnHeaders("编码").Index - 1) = !编码
'            If Me.Tag <> "7" Then
                objItem.SubItems(Me.lvwItems.ColumnHeaders("规格").Index - 1) = IIf(IsNull(!规格), "", !规格)
'            End If
            objItem.SubItems(Me.lvwItems.ColumnHeaders("产地").Index - 1) = IIf(IsNull(!产地), "", !产地)
            objItem.SubItems(Me.lvwItems.ColumnHeaders("单位").Index - 1) = IIf(IsNull(!单位), "", !单位)
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
    '功能：刷新指定药品的协定组成药品或原料药品
    '入参：lngMediId-指定的药名id
    '--------------------------------------------------------
    Err = 0: On Error GoTo ErrHand

    If Me.msfMember.Tag = "协定" Then
        gstrSql = "select I.ID,I.编码,I.名称,I.规格,I.产地,M.计算单位 as 单位,P.分子 as 采用量" & _
                " from 协定药品对照 P,收费项目目录 I,药品规格 S,诊疗项目目录 M" & _
                " where P.协定药品ID=I.ID and I.ID=S.药品ID and S.药名id=M.ID" & _
                "       and P.药品ID=[1]"
    Else
        gstrSql = "select I.ID,I.编码,I.名称,I.规格,I.产地,M.计算单位 as 单位,P.分子 as 采用量" & _
                " from 自制药品构成 P,收费项目目录 I,药品规格 S,诊疗项目目录 M" & _
                " where P.原料药品ID=I.ID and I.ID=S.药品ID and S.药名id=M.ID" & _
                "       and P.自制药品ID=[1]"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngMediId)
    
    With rsTemp
        Me.msfMember.ClearBill
        Do While Not .EOF
            If Me.msfMember.Rows < .AbsolutePosition + 1 Then Me.msfMember.Rows = Me.msfMember.Rows + 1
            Me.msfMember.RowData(.AbsolutePosition) = !ID
            Me.msfMember.TextMatrix(.AbsolutePosition, 0) = .AbsolutePosition
            Me.msfMember.TextMatrix(.AbsolutePosition, col品名) = "[" & !编码 & "]" & !名称
            Me.msfMember.TextMatrix(.AbsolutePosition, col规格) = IIf(IsNull(!规格), "", !规格)
            Me.msfMember.TextMatrix(.AbsolutePosition, col产地) = IIf(IsNull(!产地), "", !产地)
            Me.msfMember.TextMatrix(.AbsolutePosition, col采用量) = Format(!采用量, "0.00000")
            Me.msfMember.TextMatrix(.AbsolutePosition, col单位) = IIf(IsNull(!单位), "", !单位)
            .MoveNext
        Loop
    End With
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

