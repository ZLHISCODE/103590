VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.4#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmClinicLabs 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "检验项目指标"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6780
   Icon            =   "frmClinicLabs.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   6780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdRestore 
      Caption         =   "恢复(&R)"
      Height          =   350
      Left            =   2715
      Picture         =   "frmClinicLabs.frx":038A
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   4605
      Width           =   1290
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "全部清除(&C)"
      Height          =   350
      Left            =   1425
      Picture         =   "frmClinicLabs.frx":04D4
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4605
      Width           =   1290
   End
   Begin VB.TextBox txtItem 
      Height          =   300
      Left            =   1635
      MaxLength       =   50
      TabIndex        =   2
      Top             =   750
      Width           =   4620
   End
   Begin VB.CommandButton cmdItem 
      Caption         =   "&P"
      Height          =   300
      Left            =   6240
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   765
      Width           =   285
   End
   Begin MSComctlLib.ListView lvwItems 
      Height          =   2715
      Left            =   405
      TabIndex        =   8
      Top             =   5025
      Visible         =   0   'False
      Width           =   5235
      _ExtentX        =   9234
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
      Left            =   5820
      Top             =   5085
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
            Picture         =   "frmClinicLabs.frx":061E
            Key             =   "ItemUse"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "保存(&S)"
      Height          =   350
      Left            =   4335
      TabIndex        =   6
      Top             =   4605
      Width           =   1100
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "关闭(&X)"
      Height          =   350
      Left            =   5445
      TabIndex        =   7
      Top             =   4605
      Width           =   1100
   End
   Begin ZL9BillEdit.BillEdit msfLabs 
      Height          =   3090
      Left            =   225
      TabIndex        =   5
      Top             =   1425
      Width           =   6285
      _ExtentX        =   11086
      _ExtentY        =   5450
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
   Begin VB.Label lblLabs 
      AutoSize        =   -1  'True
      Caption         =   "检验标本与指标报告项目(&L)"
      Height          =   180
      Left            =   240
      TabIndex        =   4
      Top             =   1185
      Width           =   2250
   End
   Begin VB.Label lblItem 
      AutoSize        =   -1  'True
      Caption         =   "检验申请项目(&I)"
      Height          =   180
      Left            =   255
      TabIndex        =   1
      Top             =   810
      Width           =   1350
   End
   Begin VB.Label lblnote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "    检验项目可选择标本，也具有不同的检验报告指标；请选择项目后指定其可使用的标本和对应的报告指标，以便该项目执行报告的填写。"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   870
      TabIndex        =   0
      Top             =   120
      Width           =   5685
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   225
      Picture         =   "frmClinicLabs.frx":09B8
      Top             =   60
      Width           =   480
   End
End
Attribute VB_Name = "frmClinicLabs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------
'说明：
'   1、当前状态：由me.cmdClose.tag保存，分别为"修改"、"查阅"，由上级程序通过ShowMe函数传入
'   2、指定项目：由me.lblItem.tag保存，由上级程序通过ShowMe函数传入，可以传递，也可以不传递
'---------------------------------------------------
Private strInputed As String
Dim rsTemp As New ADODB.Recordset
Dim objItem As ListItem
Dim strTemp As String
Dim intCount As Integer

Public Sub ShowMe(ByVal frmParent As Object, ByVal blnEdit As Boolean, Optional ByVal lng项目id As Long)
    '---------------------------------------------------
    '功能：上级程序调用本窗体的，传递参数，并显示窗体
    '---------------------------------------------------
    Me.cmdClose.Tag = IIf(blnEdit, "修改", "查阅")
    If Me.cmdClose.Tag = "查阅" Then
        Me.msfLabs.Active = False
        Me.cmdSave.Visible = False
        Me.cmdClear.Visible = False
        Me.cmdRestore.Visible = False
    Else
        Me.msfLabs.Active = True
    End If
    Me.lblItem.Tag = lng项目id
    
    Err = 0: On Error GoTo ErrHand
    
    gstrSql = "select I.ID,I.编码,I.名称,I.计算单位,I.标本部位" & _
            " from 诊疗项目目录 I" & _
            " where I.类别='C' and I.ID=[1] " & _
            "       and (I.撤档时间 is null or I.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Me.lblItem.Tag))
    
    With rsTemp
        If .BOF Or .EOF Then
            Me.lblItem.Tag = 0: Me.txtItem.Tag = "": Me.txtItem.Text = Me.txtItem.Tag
            Me.msfLabs.Tag = ""
        Else
            Me.lblItem.Tag = !ID: Me.txtItem.Tag = "[" & !编码 & "]" & !名称: Me.txtItem.Text = Me.txtItem.Tag
            Me.msfLabs.Tag = Nvl(!标本部位)
            Call zlLabsRef(Me.lblItem.Tag)
        End If
    End With
    Me.Show 1, frmParent
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdClear_Click()
    Me.msfLabs.ClearBill
End Sub

Private Sub cmdClose_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub cmdRestore_Click()
    Call zlLabsRef(Me.lblItem.Tag)
End Sub

Private Sub cmdSave_Click()
    If Val(Me.lblItem.Tag) = 0 Then MsgBox "未正确指定诊疗项目！", vbExclamation, gstrSysName: Me.txtItem.SetFocus: Exit Sub
    strTemp = "": gstrSql = ""
    With Me.msfLabs
        For intCount = 1 To .Rows - 1
            If Trim(.TextMatrix(intCount, 1)) <> "" Then  'And Val(.TextMatrix(intCount, 2)) <> 0 Then
                If InStr(1, strTemp & ";", ";" & .RowData(intCount) & ";") > 0 Then
                    MsgBox intCount & "与前面行重复(检查标本和报告项目相同)！", vbInformation, gstrSysName
                    .SetFocus: Exit Sub
                End If
                strTemp = strTemp & ";" & Trim(.TextMatrix(intCount, 1)) & "^" & Val(.TextMatrix(intCount, 2))
                gstrSql = gstrSql & "|" & Trim(.TextMatrix(intCount, 1)) & "^" & Val(.TextMatrix(intCount, 2))
            End If
        Next
    End With
    If gstrSql <> "" Then gstrSql = Mid(gstrSql, 2)
    gstrSql = "zl_检验报告项目_UPDATE(" & Val(Me.lblItem.Tag) & ",'" & gstrSql & "')"
    
    Err = 0: On Error GoTo ErrHand
    Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
    MsgBox Me.txtItem.Text & "检验报告项目保存成功！", vbExclamation, gstrSysName
    Me.txtItem.SetFocus
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub



Private Sub cmdItem_Click()
    With Me.lvwItems.ColumnHeaders
        .Clear
        .Add , "名称", "名称", 2500
        .Add , "编码", "编码", 1000
        .Add , "计算单位", "单位", 800
    End With
    Err = 0: On Error GoTo ErrHand
    gstrSql = "select I.ID,I.编码,I.名称,I.计算单位,I.标本部位" & _
            " from 诊疗项目目录 I" & _
            " where I.类别>='A'" & _
            "       and (I.撤档时间 is null or I.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))"
'        If .State = adStateOpen Then .Close
'        Call SQLTest(App.Title, Me.Caption, gstrSql)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "cmdItem_Click")
'        Call SQLTest
    With rsTemp
        If .BOF Or .EOF Then
            MsgBox "请建立项目品种后设置配伍禁忌", vbExclamation, gstrSysName
            Me.lblItem.Tag = 0: Me.txtItem.Tag = "": Me.txtItem.Text = Me.txtItem.Tag: Me.txtItem.SetFocus: Exit Sub
        End If
        Me.lvwItems.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !名称)
            objItem.Icon = "ItemUse": objItem.SmallIcon = "ItemUse"
            objItem.SubItems(Me.lvwItems.ColumnHeaders("编码").Index - 1) = !编码
            objItem.SubItems(Me.lvwItems.ColumnHeaders("计算单位").Index - 1) = Nvl(!计算单位)
            objItem.Tag = Nvl(!标本部位)
            .MoveNext
        Loop
        Me.lvwItems.ListItems(1).Selected = True
    End With
    With Me.lvwItems
        .ColumnHeaders("编码").Position = 1
        .Width = Me.txtItem.Width
        .Tag = Me.txtItem.Name
        .Left = Me.txtItem.Left
        .Top = Me.txtItem.Top + Me.txtItem.Height
        .ZOrder 0: .Visible = True
        .SetFocus
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
        If Me.lvwItems.Tag = Me.txtItem.Name Then
            Me.txtItem.SetFocus
        Else
            Me.msfLabs.SetFocus
        End If
    Else
        cmdClose_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    With Me.msfLabs
        .Active = True
        .MsfObj.FixedCols = 1: .Rows = 2: .Cols = 6
        .TextMatrix(0, 0) = "": .TextMatrix(0, 1) = "检验标本": .TextMatrix(0, 2) = "报告项目ID": .TextMatrix(0, 3) = "报告项目": .TextMatrix(0, 4) = "类型": .TextMatrix(0, 5) = "单位"
        .ColData(0) = 5: .ColData(1) = 1: .ColData(2) = 5: .ColData(3) = 1: .ColData(4) = 5: .ColData(5) = 5
        .ColWidth(0) = 0: .ColWidth(1) = 1200: .ColWidth(2) = 0: .ColWidth(3) = 3600: .ColWidth(4) = 500: .ColWidth(5) = 600
        .TextMatrix(1, 0) = "1"
        .PrimaryCol = 1: .LocateCol = 1
        .Row = 1: .Col = 1
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
        If .Tag = Me.txtItem.Name Then
            If Me.lblItem.Tag <> Mid(.SelectedItem.Key, 2) Then
                Me.lblItem.Tag = Mid(.SelectedItem.Key, 2)
                Me.txtItem.Tag = "[" & .SelectedItem.SubItems(.ColumnHeaders("编码").Index - 1) & "]" & .SelectedItem.Text
                Me.txtItem.Text = Me.txtItem.Tag
                Me.msfLabs.Tag = .SelectedItem.Tag
                Call zlLabsRef(Me.lblItem.Tag)
            End If
            Me.txtItem.SetFocus
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf .Tag = "标本项目" Then
            Me.msfLabs.Text = .SelectedItem.Text
            Me.msfLabs.TextMatrix(Me.msfLabs.Row, 1) = Me.msfLabs.Text
            Me.msfLabs.SetFocus
            Call zlCommFun.PressKey(vbKeyRight)
        ElseIf .Tag = "报告项目" Then
            Me.msfLabs.Text = "[" & .SelectedItem.SubItems(.ColumnHeaders("编码").Index - 1) & "]" & .SelectedItem.Text
            If Trim(.SelectedItem.SubItems(.ColumnHeaders("英文名").Index - 1)) <> "" Then
                Me.msfLabs.Text = Me.msfLabs.Text & " (" & .SelectedItem.SubItems(.ColumnHeaders("英文名").Index - 1) & ")"
            End If
            Me.msfLabs.TextMatrix(Me.msfLabs.Row, 2) = Mid(.SelectedItem.Key, 2)
            Me.msfLabs.TextMatrix(Me.msfLabs.Row, 3) = Me.msfLabs.Text
            Me.msfLabs.TextMatrix(Me.msfLabs.Row, 4) = .SelectedItem.SubItems(.ColumnHeaders("类型").Index - 1)
            Me.msfLabs.TextMatrix(Me.msfLabs.Row, 5) = .SelectedItem.SubItems(.ColumnHeaders("单位").Index - 1)
            Me.msfLabs.SetFocus
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

Private Sub msfLabs_CommandClick()
    Err = 0: On Error GoTo ErrHand
    If Me.msfLabs.Col = 1 Then
        With Me.lvwItems.ColumnHeaders
            .Clear
            .Add , "名称", "名称", 1700
            .Add , "编码", "编码", 500
            .Add , "简码", "简码", 800
        End With
        
        gstrSql = "select 编码,名称,简码 from 诊疗检验标本 order by 编码"
'            If .State = adStateOpen Then .Close
'            Call SQLTest(App.ProductName, Me.Caption, gstrSql)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "msfLabs_CommandClick")
'            Call SQLTest
        With rsTemp
            If .BOF Or .EOF Then
                MsgBox "请建立检验标本后进行(字典管理)！", vbExclamation, gstrSysName: Exit Sub
            End If
            Me.lvwItems.ListItems.Clear
            Do While Not .EOF
                Set objItem = Me.lvwItems.ListItems.Add(, "_" & !编码, !名称)
                objItem.Icon = "ItemUse": objItem.SmallIcon = "ItemUse"
                objItem.SubItems(Me.lvwItems.ColumnHeaders("编码").Index - 1) = !编码
                objItem.SubItems(Me.lvwItems.ColumnHeaders("简码").Index - 1) = Nvl(!简码)
                .MoveNext
            Loop
            Me.lvwItems.ListItems(1).Selected = True
        End With
        With Me.lvwItems
            .ColumnHeaders("编码").Position = 1
            .Width = 3600
            .Tag = "标本项目"
            .Left = Me.msfLabs.Left + 300
            .Top = Me.msfLabs.Top
            .ZOrder 0: .Visible = True
            .SetFocus
        End With
    Else
        With Me.lvwItems.ColumnHeaders
            .Clear
            .Add , "中文名", "中文名", 1600
            .Add , "编码", "编码", 900
            .Add , "英文名", "英文名", 1100
            .Add , "类型", "类型", 600
            .Add , "单位", "单位", 700
        End With
        
        gstrSql = "select I.ID,I.编码,I.中文名,I.英文名,I.类型,I.单位" & _
                " from 诊治所见分类 C,诊治所见项目 I" & _
                " where C.ID=I.分类ID and C.性质=3" & _
                " order by I.编码"
'            If .State = adStateOpen Then .Close
'            Call SQLTest(App.ProductName, Me.Caption, gstrSql)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "msfLabs_CommandClick")
'            Call SQLTest
        With rsTemp
            If .BOF Or .EOF Then
                MsgBox "尚未建立检验所见报告项目！", vbExclamation, gstrSysName: Exit Sub
            End If
            Me.lvwItems.ListItems.Clear
            Do While Not .EOF
                Set objItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !中文名)
                objItem.Icon = "ItemUse": objItem.SmallIcon = "ItemUse"
                objItem.SubItems(Me.lvwItems.ColumnHeaders("编码").Index - 1) = !编码
                objItem.SubItems(Me.lvwItems.ColumnHeaders("英文名").Index - 1) = Nvl(!英文名)
                Select Case Nvl(!类型, 0)
                Case 0
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("类型").Index - 1) = "数值"
                Case 1
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("类型").Index - 1) = "文字"
                Case 2
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("类型").Index - 1) = "日期"
                Case 3
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("类型").Index - 1) = "逻辑"
                End Select
                objItem.SubItems(Me.lvwItems.ColumnHeaders("单位").Index - 1) = Nvl(!单位)
                .MoveNext
            Loop
            Me.lvwItems.ListItems(1).Selected = True
        End With
        With Me.lvwItems
            .ColumnHeaders("编码").Position = 1
            .Width = 5200
            .Tag = "报告项目"
            .Left = Me.msfLabs.Left + 1200
            .Top = Me.msfLabs.Top
            .ZOrder 0: .Visible = True
            .SetFocus
        End With
    End If
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub msfLabs_EnterCell(Row As Long, Col As Long)
    strInputed = Me.msfLabs.TextMatrix(Row, Col)
End Sub

Private Sub msfLabs_GotFocus()
    If Me.lvwItems.Visible Then Me.lvwItems.SetFocus
End Sub

Private Sub msfLabs_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    If KeyCode <> vbKeyReturn Then Exit Sub
    With Me.msfLabs
        If .Active = False Then Exit Sub
        If .TxtVisible = False Then
            If .TextMatrix(.Row, 1) = "" Then .SetFocus: KeyCode = 0: Call zlCommFun.PressKey(vbKeyTab): Exit Sub
            strTemp = UCase(Trim(.TextMatrix(.Row, .Col)))
        Else
            If Trim(.Text) = "" Then .SetFocus: KeyCode = 0: Call zlCommFun.PressKey(vbKeyTab): Exit Sub
            strTemp = UCase(Trim(.Text))
        End If
    End With
    If strTemp = strInputed Then Exit Sub
    If InStr(1, strTemp, "[") <> 0 And InStr(1, strTemp, "]") <> 0 Then strTemp = Mid(strTemp, 2, InStr(1, strTemp, "]") - 2)
    
    Err = 0: On Error GoTo ErrHand
    If Me.msfLabs.Col = 1 Then
        With Me.lvwItems.ColumnHeaders
            .Clear
            .Add , "名称", "名称", 1700
            .Add , "编码", "编码", 500
            .Add , "简码", "简码", 800
        End With
        
        gstrSql = "select 编码,名称,简码" & _
                " from 诊疗检验标本" & _
                " where (编码 like [1] or 名称 like [2] or 简码 like [2])"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, strTemp & "%", gstrMatch & strTemp & "%")
        
        With rsTemp
            If .BOF Or .EOF Then
                MsgBox "未找到指定标本，请重新输入！", vbExclamation, gstrSysName: Cancel = True: Exit Sub
            End If
            If .RecordCount = 1 Then
                Me.msfLabs.Text = !名称
                Me.msfLabs.TextMatrix(Me.msfLabs.Row, 1) = Me.msfLabs.Text
                Exit Sub
            End If
            Me.lvwItems.ListItems.Clear
            Do While Not .EOF
                Set objItem = Me.lvwItems.ListItems.Add(, "_" & !编码, !名称)
                objItem.Icon = "ItemUse": objItem.SmallIcon = "ItemUse"
                objItem.SubItems(Me.lvwItems.ColumnHeaders("编码").Index - 1) = !编码
                objItem.SubItems(Me.lvwItems.ColumnHeaders("简码").Index - 1) = Nvl(!简码)
                .MoveNext
            Loop
            Me.lvwItems.ListItems(1).Selected = True
        End With
        With Me.lvwItems
            .ColumnHeaders("编码").Position = 1
            .Width = 3200
            .Tag = "标本项目"
            .Left = Me.msfLabs.Left + 300
            .Top = Me.msfLabs.Top
            .ZOrder 0: .Visible = True
            .SetFocus
        End With
    Else
        With Me.lvwItems.ColumnHeaders
            .Clear
            .Add , "中文名", "中文名", 1600
            .Add , "编码", "编码", 900
            .Add , "英文名", "英文名", 1100
            .Add , "类型", "类型", 600
            .Add , "单位", "单位", 700
        End With
        
        gstrSql = "select I.ID,I.编码,I.中文名,I.英文名,I.类型,I.单位" & _
                " from 诊治所见分类 C,诊治所见项目 I" & _
                " where C.ID=I.分类ID and C.性质=3" & _
                "       and (I.编码 like [1] or I.中文名 like [2] or upper(I.英文名) like [2])" & _
                " order by I.编码"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, strTemp & "%", gstrMatch & strTemp & "%")
        
        With rsTemp
            If .BOF Or .EOF Then
                MsgBox "未找到指定的检验所见报告项目！", vbExclamation, gstrSysName: Cancel = True: Exit Sub
            End If
            If .RecordCount = 1 Then
                Me.msfLabs.Text = "[" & !编码 & "]" & !中文名 & IIf(IsNull(!英文名), "", " (" & !英文名 & ")")
                Me.msfLabs.TextMatrix(Me.msfLabs.Row, 2) = !ID
                Me.msfLabs.TextMatrix(Me.msfLabs.Row, 3) = Me.msfLabs.Text
                Select Case IIf(IsNull(!类型), 0, !类型)
                Case 0
                    Me.msfLabs.TextMatrix(Me.msfLabs.Row, 4) = "数值"
                Case 1
                    Me.msfLabs.TextMatrix(Me.msfLabs.Row, 4) = "文字"
                Case 2
                    Me.msfLabs.TextMatrix(Me.msfLabs.Row, 4) = "日期"
                Case 3
                    Me.msfLabs.TextMatrix(Me.msfLabs.Row, 4) = "逻辑"
                End Select
                Me.msfLabs.TextMatrix(Me.msfLabs.Row, 5) = Nvl(!单位)
                Exit Sub
            End If
            Me.lvwItems.ListItems.Clear
            Do While Not .EOF
                Set objItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !中文名)
                objItem.Icon = "ItemUse": objItem.SmallIcon = "ItemUse"
                objItem.SubItems(Me.lvwItems.ColumnHeaders("编码").Index - 1) = !编码
                objItem.SubItems(Me.lvwItems.ColumnHeaders("英文名").Index - 1) = Nvl(!英文名)
                Select Case Nvl(!类型, 0)
                Case 0
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("类型").Index - 1) = "数值"
                Case 1
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("类型").Index - 1) = "文字"
                Case 2
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("类型").Index - 1) = "日期"
                Case 3
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("类型").Index - 1) = "逻辑"
                End Select
                objItem.SubItems(Me.lvwItems.ColumnHeaders("单位").Index - 1) = Nvl(!单位)
                .MoveNext
            Loop
            Me.lvwItems.ListItems(1).Selected = True
        End With
        With Me.lvwItems
            .ColumnHeaders("编码").Position = 1
            .Width = 5200
            .Tag = "报告项目"
            .Left = Me.msfLabs.Left + 1200
            .Top = Me.msfLabs.Top
            .ZOrder 0: .Visible = True
            .SetFocus
        End With
    End If
    Cancel = True
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtItem_GotFocus()
    Me.txtItem.SelStart = 0: Me.txtItem.SelLength = 100
End Sub

Private Sub txtItem_KeyPress(KeyAscii As Integer)
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If KeyAscii <> vbKeyReturn Then Exit Sub
    strTemp = UCase(Trim(Me.txtItem.Text))
    If strTemp = "" Then Me.lblItem.Tag = 0: Me.txtItem.Tag = "": Me.txtItem.Text = "": Exit Sub
    
    If InStr(1, strTemp, "[") <> 0 And InStr(1, strTemp, "]") <> 0 Then strTemp = Mid(strTemp, 2, InStr(1, strTemp, "]") - 2)
    With Me.lvwItems.ColumnHeaders
        .Clear
        .Add , "名称", "名称", 2500
        .Add , "编码", "编码", 1000
        .Add , "计算单位", "单位", 800
    End With
    Err = 0: On Error GoTo ErrHand
    
    gstrSql = "select distinct I.ID,I.编码,I.名称,I.计算单位,I.标本部位" & _
            " from 诊疗项目目录 I,诊疗项目别名 N" & _
            " where I.ID=N.诊疗项目ID and I.类别>='A'" & _
            "       and (I.撤档时间 is null or I.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))" & _
            "       and (I.编码 like [1] or N.名称 like [2] or N.简码 like [2])"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, strTemp & "%", gstrMatch & strTemp & "%")
        
    With rsTemp
        If .BOF Or .EOF Then
            MsgBox "未找到指定的诊疗项目，请重新指定", vbExclamation, gstrSysName
            Me.lblItem.Tag = 0: Me.txtItem.Tag = "": Me.txtItem.Text = Me.txtItem.Tag: Me.txtItem.SetFocus: Exit Sub
        End If
        If .RecordCount = 1 Then
            If Me.lblItem.Tag <> !ID Then
                Me.lblItem.Tag = !ID: Me.txtItem.Tag = "[" & !编码 & "]" & !名称: Me.txtItem.Text = Me.txtItem.Tag
                Me.msfLabs.Tag = Nvl(!标本部位)
                Call zlLabsRef(Me.lblItem.Tag)
            End If
            Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        Me.lvwItems.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !名称)
            objItem.Icon = "ItemUse": objItem.SmallIcon = "ItemUse"
            objItem.SubItems(Me.lvwItems.ColumnHeaders("编码").Index - 1) = !编码
            objItem.SubItems(Me.lvwItems.ColumnHeaders("计算单位").Index - 1) = Nvl(!计算单位)
            objItem.Tag = Nvl(!标本部位)
            .MoveNext
        Loop
        Me.lvwItems.ListItems(1).Selected = True
    End With
    With Me.lvwItems
        .ColumnHeaders("编码").Position = 1
        .Width = Me.txtItem.Width
        .Tag = Me.txtItem.Name
        .Left = Me.txtItem.Left
        .Top = Me.txtItem.Top + Me.txtItem.Height
        .ZOrder 0: .Visible = True
        .SetFocus
    End With
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtItem_LostFocus()
    Me.txtItem.Text = Me.txtItem.Tag
End Sub

Private Sub zlLabsRef(lngItemID As Long)
    '--------------------------------------------------------
    '功能：刷新显示诊疗项目对应的报告项目内容
    '入参：lngItemId-指定的诊疗项目id(此处为检验类项目)
    '--------------------------------------------------------
    Err = 0: On Error GoTo ErrHand
    
    gstrSql = "select R.检验标本,R.报告项目ID,'['||I.编码||']'||I.中文名||decode(I.英文名,null,'',' ('||I.英文名||')') as 报告项目,I.类型,I.单位" & _
            " from 检验报告项目 R,诊治所见项目 I" & _
            " where R.报告项目ID=I.ID(+) and R.诊疗项目ID=[1] " & _
            " order by R.检验标本,R.排列序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemID)
    
    With rsTemp
        Me.msfLabs.ClearBill
        Do While Not .EOF
            If Me.msfLabs.Rows - 1 < .AbsolutePosition Then Me.msfLabs.Rows = Me.msfLabs.Rows + 1
            Me.msfLabs.TextMatrix(.AbsolutePosition, 0) = .AbsolutePosition
            Me.msfLabs.TextMatrix(.AbsolutePosition, 1) = !检验标本
            Me.msfLabs.TextMatrix(.AbsolutePosition, 2) = Nvl(!报告项目ID, 0)
            If Me.msfLabs.TextMatrix(.AbsolutePosition, 2) = "0" Then
                Me.msfLabs.TextMatrix(.AbsolutePosition, 3) = ""
            Else
                Me.msfLabs.TextMatrix(.AbsolutePosition, 3) = Nvl(!报告项目)
            End If
            Select Case IIf(IsNull(!类型), 0, !类型)
            Case 0
                Me.msfLabs.TextMatrix(.AbsolutePosition, 4) = "数值"
            Case 1
                Me.msfLabs.TextMatrix(.AbsolutePosition, 4) = "文字"
            Case 2
                Me.msfLabs.TextMatrix(.AbsolutePosition, 4) = "日期"
            Case 3
                Me.msfLabs.TextMatrix(.AbsolutePosition, 4) = "逻辑"
            End Select
            If Me.msfLabs.TextMatrix(.AbsolutePosition, 2) = "0" Then
                Me.msfLabs.TextMatrix(.AbsolutePosition, 4) = ""
            End If
            Me.msfLabs.TextMatrix(.AbsolutePosition, 5) = Nvl(!单位)
            .MoveNext
        Loop
    End With
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


