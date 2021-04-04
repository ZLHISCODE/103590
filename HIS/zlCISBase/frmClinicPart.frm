VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.4#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmClinicPart 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "检查部位组合"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6780
   Icon            =   "frmClinicPart.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   6780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdRestore 
      Caption         =   "恢复(&R)"
      Height          =   350
      Left            =   2715
      Picture         =   "frmClinicPart.frx":08CA
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   4455
      Width           =   1290
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "全部清除(&C)"
      Height          =   350
      Left            =   1425
      Picture         =   "frmClinicPart.frx":0A14
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   4455
      Width           =   1290
   End
   Begin VB.TextBox txtItem 
      Height          =   300
      Left            =   1830
      MaxLength       =   50
      TabIndex        =   2
      Top             =   810
      Width           =   4425
   End
   Begin VB.CommandButton cmdItem 
      Caption         =   "&P"
      Height          =   300
      Left            =   6240
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   825
      Width           =   285
   End
   Begin MSComctlLib.ListView lvwItems 
      Height          =   2715
      Left            =   405
      TabIndex        =   9
      Top             =   4815
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
      Left            =   5835
      Top             =   4875
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
            Picture         =   "frmClinicPart.frx":0B5E
            Key             =   "ItemUse"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "保存(&S)"
      Height          =   350
      Left            =   4335
      TabIndex        =   6
      Top             =   4455
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   225
      Picture         =   "frmClinicPart.frx":10F8
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4455
      Width           =   1100
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "关闭(&X)"
      Height          =   350
      Left            =   5445
      TabIndex        =   7
      Top             =   4455
      Width           =   1100
   End
   Begin ZL9BillEdit.BillEdit msfPart 
      Height          =   2880
      Left            =   225
      TabIndex        =   5
      Top             =   1500
      Width           =   6285
      _ExtentX        =   11086
      _ExtentY        =   5080
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
   Begin VB.Label lblPart 
      AutoSize        =   -1  'True
      Caption         =   "同名单部位项目(&E)"
      Height          =   180
      Left            =   240
      TabIndex        =   4
      Top             =   1260
      Width           =   1530
   End
   Begin VB.Label lblItem 
      AutoSize        =   -1  'True
      Caption         =   "多部位可选项目(&I)"
      Height          =   180
      Left            =   255
      TabIndex        =   1
      Top             =   870
      Width           =   1530
   End
   Begin VB.Label lblnote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmClinicPart.frx":1242
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   870
      TabIndex        =   0
      Top             =   120
      Width           =   5685
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   225
      Picture         =   "frmClinicPart.frx":12D0
      Top             =   90
      Width           =   480
   End
End
Attribute VB_Name = "frmClinicPart"
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

Public Sub ShowME(ByVal frmParent As Object, ByVal blnEdit As Boolean, Optional ByVal lng项目id As Long)
    '---------------------------------------------------
    '功能：上级程序调用本窗体的，传递参数，并显示窗体
    '---------------------------------------------------
    Me.cmdClose.Tag = IIf(blnEdit, "修改", "查阅")
    If Me.cmdClose.Tag = "查阅" Then
        Me.msfPart.Active = False
        Me.cmdSave.Visible = False
        Me.cmdClear.Visible = False
        Me.cmdRestore.Visible = False
    Else
        Me.msfPart.Active = True
    End If
    Me.lblItem.Tag = lng项目id
    
    Err = 0: On Error GoTo ErrHand
    
    gstrSql = "select I.ID,I.编码,I.名称,I.计算单位" & _
            " from 诊疗项目目录 I" & _
            " where I.类别='D' and I.组合项目=1 and I.ID=[1] " & _
            "       and (I.撤档时间 is null or I.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Me.lblItem.Tag))
    
    With rsTemp
        If .BOF Or .EOF Then
            Me.lblItem.Tag = 0: Me.txtItem.Tag = "": Me.txtItem.Text = Me.txtItem.Tag
        Else
            Me.lblItem.Tag = !ID: Me.txtItem.Tag = "[" & !编码 & "]" & !名称: Me.txtItem.Text = Me.txtItem.Tag
            Call zlPartRef(Me.lblItem.Tag)
        End If
    End With
    Me.Show 1, frmParent
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdClear_Click()
    Me.msfPart.ClearBill
End Sub

Private Sub cmdClose_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub cmdRestore_Click()
    Call zlPartRef(Me.lblItem.Tag)
End Sub

Private Sub cmdSave_Click()
    If Val(Me.lblItem.Tag) = 0 Then MsgBox "未正确指定诊疗项目！", vbExclamation, gstrSysName: Me.txtItem.SetFocus: Exit Sub
    strTemp = "": gstrSql = ""
    With Me.msfPart
        For intCount = 1 To .Rows - 1
            If Trim(.TextMatrix(intCount, 1)) <> "" And .RowData(intCount) <> 0 Then
                If InStr(1, strTemp & ";", ";" & .RowData(intCount) & ";") > 0 Then
                    MsgBox intCount & "行项目与前面项目有重复！", vbInformation, gstrSysName
                    .SetFocus: Exit Sub
                End If
                strTemp = strTemp & ";" & .RowData(intCount)
                gstrSql = gstrSql & "|" & .RowData(intCount)
            End If
        Next
    End With
    If gstrSql <> "" Then gstrSql = Mid(gstrSql, 2)
    gstrSql = "zl_检查组合项目_UPDATE(" & Val(Me.lblItem.Tag) & ",'" & gstrSql & "')"
    
    Err = 0: On Error GoTo ErrHand
    Call zldatabase.ExecuteProcedure(gstrSql, Me.Caption)
    MsgBox Me.txtItem.Text & " 部位组合保存成功！", vbExclamation, gstrSysName
    Me.txtItem.SetFocus
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdItem_Click()
    Err = 0: On Error GoTo ErrHand
    gstrSql = "select I.ID,I.编码,I.名称,'(可选...)' as 标本部位" & _
            " from 诊疗项目目录 I" & _
            " where I.类别='D' and I.组合项目=1" & _
            "       and (I.撤档时间 is null or I.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))"
'        If .State = adStateOpen Then .Close
'        Call SQLTest(App.Title, Me.Caption, gstrSql)
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "cmdItem_Click")
'        Call SQLTest
    With rsTemp
        If .BOF Or .EOF Then
            MsgBox "目前不存在多部位组合项目", vbExclamation, gstrSysName
            Me.lblItem.Tag = 0: Me.txtItem.Tag = "": Me.txtItem.Text = Me.txtItem.Tag: Me.txtItem.SetFocus: Exit Sub
        End If
        Me.lvwItems.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !名称)
            objItem.Icon = "ItemUse": objItem.SmallIcon = "ItemUse"
            objItem.SubItems(Me.lvwItems.ColumnHeaders("编码").Index - 1) = !编码
            objItem.SubItems(Me.lvwItems.ColumnHeaders("标本部位").Index - 1) = !标本部位
            .MoveNext
        Loop
        Me.lvwItems.ListItems(1).Selected = True
    End With
    With Me.lvwItems
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
            Me.msfPart.SetFocus
        End If
    Else
        cmdClose_Click
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    With Me.msfPart
        .Active = True
        .MsfObj.FixedCols = 1: .Rows = 2: .Cols = 3
        .TextMatrix(0, 0) = "": .TextMatrix(0, 1) = "检查项目": .TextMatrix(0, 2) = "检查部位"
        .ColData(0) = 5: .ColData(1) = 1: .ColData(2) = 5
        .ColWidth(0) = 250: .ColWidth(1) = 3500: .ColWidth(2) = 2000
        .TextMatrix(1, 0) = "1"
        .PrimaryCol = 1: .LocateCol = 1
        .Row = 1: .Col = 1
    End With
    
    Me.lvwItems.ListItems.Clear
    With Me.lvwItems.ColumnHeaders
        .Clear
        .Add , "名称", "名称", 2500
        .Add , "编码", "编码", 1000
        .Add , "标本部位", "部位", 1600
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
        If .Tag = Me.txtItem.Name Then
            If Me.lblItem.Tag <> Mid(.SelectedItem.Key, 2) Then
                Me.lblItem.Tag = Mid(.SelectedItem.Key, 2)
                Me.txtItem.Tag = "[" & .SelectedItem.SubItems(.ColumnHeaders("编码").Index - 1) & "]" & .SelectedItem.Text
                Me.txtItem.Text = Me.txtItem.Tag
                Call zlPartRef(Me.lblItem.Tag)
            End If
            Me.txtItem.SetFocus
            Call zlcommfun.PressKey(vbKeyTab)
        Else
            Me.msfPart.Text = "[" & .SelectedItem.SubItems(.ColumnHeaders("编码").Index - 1) & "]" & .SelectedItem.Text
            Me.msfPart.RowData(Me.msfPart.Row) = Mid(.SelectedItem.Key, 2)
            Me.msfPart.TextMatrix(Me.msfPart.Row, 1) = Me.msfPart.Text
            Me.msfPart.TextMatrix(Me.msfPart.Row, 2) = .SelectedItem.SubItems(.ColumnHeaders("标本部位").Index - 1)
            Me.msfPart.SetFocus
            Call zlcommfun.PressKey(vbKeyRight)
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

Private Sub msfPart_AfterAddRow(Row As Long)
    With Me.msfPart
        For intCount = Row To .Rows - 1
            .TextMatrix(intCount, 0) = intCount
        Next
    End With
End Sub

Private Sub msfPart_AfterDeleteRow()
    With Me.msfPart
        For intCount = IIf(.Row <> 1, .Row - 1, .Row) To .Rows - 1
            .TextMatrix(intCount, 0) = intCount
        Next
    End With
End Sub

Private Sub msfPart_CommandClick()
    Err = 0: On Error GoTo ErrHand
    
    gstrSql = "select I.ID,I.编码,I.名称,I.标本部位" & _
            " from 诊疗项目目录 I" & _
            " where I.类别='D' and nvl(I.组合项目,0)=0 and I.标本部位 is not null" & _
            "       and I.名称=[1] " & _
            "       and (I.撤档时间 is null or I.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, Trim(Mid(Me.txtItem.Text, InStr(1, Me.txtItem.Text, "]") + 1)))
    
    With rsTemp
        If .BOF Or .EOF Then
            MsgBox "目前没有建立同名的单部位检查项目，无法设置！", vbExclamation, gstrSysName: Exit Sub
        End If
        Me.lvwItems.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !名称)
            objItem.Icon = "ItemUse": objItem.SmallIcon = "ItemUse"
            objItem.SubItems(Me.lvwItems.ColumnHeaders("编码").Index - 1) = !编码
            objItem.SubItems(Me.lvwItems.ColumnHeaders("标本部位").Index - 1) = !标本部位
            .MoveNext
        Loop
        Me.lvwItems.ListItems(1).Selected = True
    End With
    With Me.lvwItems
        .Tag = Me.msfPart.Name
        .Left = Me.msfPart.Left + 300
        .Top = Me.msfPart.Top
        .ZOrder 0: .Visible = True
        .SetFocus
    End With
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub msfPart_EnterCell(Row As Long, Col As Long)
    strInputed = Me.msfPart.TextMatrix(Row, Col)
End Sub

Private Sub msfPart_GotFocus()
    If Me.lvwItems.Visible Then Me.lvwItems.SetFocus
End Sub

Private Sub msfPart_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    If KeyCode <> vbKeyReturn Then Exit Sub
    With Me.msfPart
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
    If strTemp = strInputed Then Exit Sub
    If InStr(1, strTemp, "[") <> 0 And InStr(1, strTemp, "]") <> 0 Then strTemp = Mid(strTemp, 2, InStr(1, strTemp, "]") - 2)
    
    Err = 0: On Error GoTo ErrHand
    
    gstrSql = "select distinct I.ID,I.编码,I.名称,I.标本部位" & _
            " from 诊疗项目目录 I,诊疗项目别名 N" & _
            " where I.ID=N.诊疗项目id and I.类别='D' and nvl(I.组合项目,0)=0 and I.标本部位 is not null" & _
            "       and I.名称=[1] " & _
            "       and (I.撤档时间 is null or I.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))" & _
            "       and (I.编码 like [2] or N.名称 like [3] or N.简码 like [3])"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, Trim(Mid(Me.txtItem.Text, InStr(1, Me.txtItem.Text, "]") + 1)), strTemp & "%", gstrMatch & strTemp & "%")
    
    With rsTemp
        If .BOF Or .EOF Then
            MsgBox "未找到同名的单部位检查项目，请重新输入！", vbExclamation, gstrSysName: Cancel = True: Exit Sub
        End If
        If .RecordCount = 1 Then
            Me.msfPart.Text = "[" & !编码 & "]" & !名称
            Me.msfPart.RowData(Me.msfPart.Row) = !ID
            Me.msfPart.TextMatrix(Me.msfPart.Row, 1) = Me.msfPart.Text
            Me.msfPart.TextMatrix(Me.msfPart.Row, 2) = !标本部位
            Exit Sub
        End If
        Me.lvwItems.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !名称)
            objItem.Icon = "ItemUse": objItem.SmallIcon = "ItemUse"
            objItem.SubItems(Me.lvwItems.ColumnHeaders("编码").Index - 1) = !编码
            objItem.SubItems(Me.lvwItems.ColumnHeaders("标本部位").Index - 1) = !标本部位
            .MoveNext
        Loop
        Me.lvwItems.ListItems(1).Selected = True
    End With
    With Me.lvwItems
        .Tag = Me.msfPart.Name
        .Left = Me.msfPart.Left + 300
        .Top = Me.msfPart.Top
        .ZOrder 0: .Visible = True
        .SetFocus
    End With
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
    Err = 0: On Error GoTo ErrHand
    
    gstrSql = "select distinct I.ID,I.编码,I.名称,'(可选...)' as 标本部位" & _
            " from 诊疗项目目录 I,诊疗项目别名 N" & _
            " where I.ID=N.诊疗项目ID and I.类别='D' and I.组合项目=1" & _
            "       and (I.撤档时间 is null or I.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))" & _
            "       and (I.编码 like [1] or N.名称 like [2] or N.简码 like [2])"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, strTemp & "%", gstrMatch & strTemp & "%")
        
    With rsTemp
        If .BOF Or .EOF Then
            MsgBox "未找到指定的可选部位检查项目，请重新指定", vbExclamation, gstrSysName
            Me.lblItem.Tag = 0: Me.txtItem.Tag = "": Me.txtItem.Text = Me.txtItem.Tag: Me.txtItem.SetFocus: Exit Sub
        End If
        If .RecordCount = 1 Then
            If Me.lblItem.Tag <> !ID Then
                Me.lblItem.Tag = !ID: Me.txtItem.Tag = "[" & !编码 & "]" & !名称: Me.txtItem.Text = Me.txtItem.Tag
                Call zlPartRef(Me.lblItem.Tag)
            End If
            Call zlcommfun.PressKey(vbKeyTab)
            Exit Sub
        End If
        Me.lvwItems.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !名称)
            objItem.Icon = "ItemUse": objItem.SmallIcon = "ItemUse"
            objItem.SubItems(Me.lvwItems.ColumnHeaders("编码").Index - 1) = !编码
            objItem.SubItems(Me.lvwItems.ColumnHeaders("标本部位").Index - 1) = !标本部位
            .MoveNext
        Loop
        Me.lvwItems.ListItems(1).Selected = True
    End With
    With Me.lvwItems
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

Private Sub zlPartRef(lngItemId As Long)
    '--------------------------------------------------------
    '功能：刷新显示诊疗项目对应的单部位检查项目组合
    '入参：lngItemId-指定的诊疗项目id（此处为可选多部位检查项目）
    '--------------------------------------------------------
    Err = 0: On Error GoTo ErrHand
    
    gstrSql = "select I.ID,'['||I.编码||']'||I.名称 as 名称,I.标本部位" & _
            " from 诊疗项目组合 R,诊疗项目目录 I" & _
            " where R.诊疗项目ID=I.ID and R.诊疗组合ID=[1] " & _
            " order by R.序号"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemId)
        
    With rsTemp
        Me.msfPart.ClearBill
        Do While Not .EOF
            If Me.msfPart.Rows - 1 < .AbsolutePosition Then Me.msfPart.Rows = Me.msfPart.Rows + 1
            Me.msfPart.TextMatrix(.AbsolutePosition, 0) = .AbsolutePosition
            Me.msfPart.RowData(.AbsolutePosition) = !ID
            Me.msfPart.TextMatrix(.AbsolutePosition, 1) = !名称
            Me.msfPart.TextMatrix(.AbsolutePosition, 2) = !标本部位
            .MoveNext
        Loop
    End With
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


