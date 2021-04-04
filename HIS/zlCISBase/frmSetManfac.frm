VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSetManfac 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "厂家批准文号"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6960
   Icon            =   "frmSetManfac.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   6960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdMedi 
      Caption         =   "…"
      Height          =   285
      Left            =   6600
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   698
      Width           =   285
   End
   Begin VB.TextBox txtMedi 
      Height          =   300
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   0
      Top             =   690
      Width           =   5445
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "关闭(&C)"
      Height          =   350
      Left            =   5790
      TabIndex        =   6
      Top             =   4860
      Width           =   1100
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "删除(&D)"
      Height          =   350
      Left            =   0
      TabIndex        =   4
      Top             =   4860
      Width           =   1100
   End
   Begin VB.CommandButton cmdRestore 
      Caption         =   "恢复(&R)"
      Height          =   350
      Left            =   1090
      TabIndex        =   5
      Top             =   4860
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "保存(&S)"
      Height          =   350
      Left            =   4710
      TabIndex        =   3
      Top             =   4860
      Width           =   1095
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfUnit 
      Height          =   3375
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   6735
      _cx             =   11880
      _cy             =   5953
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   10329501
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   3
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSetManfac.frx":6852
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   2
      ShowComboButton =   1
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin MSComctlLib.ListView lvwItems 
      Height          =   2790
      Left            =   120
      TabIndex        =   10
      Top             =   5520
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
      Left            =   5760
      Top             =   5520
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
            Picture         =   "frmSetManfac.frx":68CD
            Key             =   "ItemUse"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetManfac.frx":6E67
            Key             =   "ItemStop"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblSave 
      Caption         =   "保存成功！"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   3600
      TabIndex        =   11
      Top             =   4950
      Width           =   975
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   120
      Picture         =   "frmSetManfac.frx":7401
      Top             =   0
      Width           =   480
   End
   Begin VB.Label lblMedi 
      AutoSize        =   -1  'True
      Caption         =   "药品规格(&M)"
      Height          =   180
      Left            =   120
      TabIndex        =   9
      Top             =   750
      Width           =   990
   End
   Begin VB.Label lblnote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "    请选择药品后，指定其生产商、批准文号。药品入库时，自动填写生产商和批准文号"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   720
      TabIndex        =   8
      Top             =   60
      Width           =   5685
   End
   Begin VB.Label lblSpec 
      AutoSize        =   -1  'True
      Caption         =   "规格：       单位：瓶"
      Height          =   180
      Left            =   120
      TabIndex        =   7
      Top             =   1200
      Width           =   1890
   End
End
Attribute VB_Name = "frmSetManfac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mObjItem As ListItem
Private mblnStar As Boolean
Private mstr类型 As String
Private mstrPrivs As String
Private mlng药品ID As String
Private mstr原值 As String
Private mblnSave As Boolean     '记录是否点击了保存按钮

Public Sub ShowMe(ByVal str类型 As String, ByVal strPrivs As String, ByVal lng药品ID As Long)
    mstr类型 = str类型
    mstrPrivs = strPrivs
    mlng药品ID = lng药品ID
    
    Me.Show vbModal, frmMediLists
End Sub

Private Sub cmdClose_Click()
    Dim intRow As Integer
    Dim intCol As Integer
    Dim strTemp As String
    
    If mblnSave = False Then
        strTemp = ""
        With vsfUnit
            For intRow = 1 To .Rows - 1
                For intCol = 1 To .Cols - 1
                    strTemp = strTemp & .TextMatrix(intRow, intCol) & "|"
                Next
            Next
        End With
        If strTemp <> mstr原值 Then
            If MsgBox("当前内容被修改后未保存，你确定要继续吗？", vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                Unload Me
            End If
        Else
            Unload Me
        End If
    Else
        Unload Me
    End If
End Sub

Private Sub cmdDel_Click()
    Dim i As Integer
    With vsfUnit
        If .Rows = 2 Then   '如果行数只有一行，则直接删除文字 ，行保留
            .TextMatrix(1, 0) = ""
            .TextMatrix(1, 1) = ""
            .TextMatrix(1, 2) = ""
            Exit Sub
        End If
        .RemoveItem (.Row) '删除当前选中行
        For i = 1 To .Rows - 1
            .TextMatrix(i, 0) = i
        Next
    End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If lvwItems.Visible = True Then
            lvwItems.Visible = False: txtMedi.SetFocus: Exit Sub
        End If
    End If
End Sub

Private Sub cmdMedi_Click()
    Dim rsTemp As ADODB.Recordset
    
    err = 0: On Error GoTo ErrHand
    
    gstrSql = "select I.ID,I.编码,I.名称,I.规格,I.产地,I.计算单位 as 单位" & _
            " from 收费项目目录 I" & _
            " where I.类别=[1] " & _
            "       and (I.撤档时间 is null or I.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Me.Tag)
    
    With rsTemp
        If .BOF Or .EOF = 1 Then
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
                        "   单位：" & IIf(IsNull(!单位), "", !单位)
                Else
                    Me.lblSpec.Caption = "   单位：" & IIf(IsNull(!单位), "", !单位)
                End If
            End If
            Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        Me.lvwItems.ListItems.Clear
        Do While Not .EOF
            Set mObjItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !名称)
            mObjItem.Icon = "ItemUse": mObjItem.SmallIcon = "ItemUse"
            mObjItem.SubItems(Me.lvwItems.ColumnHeaders("编码").Index - 1) = !编码
            mObjItem.SubItems(Me.lvwItems.ColumnHeaders("规格").Index - 1) = IIf(IsNull(!规格), "", !规格)
            mObjItem.SubItems(Me.lvwItems.ColumnHeaders("产地").Index - 1) = IIf(IsNull(!产地), "", !产地)
            mObjItem.SubItems(Me.lvwItems.ColumnHeaders("单位").Index - 1) = IIf(IsNull(!单位), "", !单位)
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

'Private Sub cmdPro_Click()
'    Dim rsProvider As New Recordset
'    Dim vRect As RECT, blnCancel As Boolean
'    Dim strPro As String
'
'    On Error Resume Next
'
'    vRect = GetControlRect(txtProInput.hWnd)
'
'    gstrSql = "select rownum as id, 名称 from 药品生产商 where 名称 like [1] or 简码 like [1] or 编码 like [1]"
'    Set rsProvider = zlDatabase.ShowSQLSelect(Me, gstrSql, 0, "产地", False, "", "", False, False, _
'    True, vRect.Left, vRect.Top, txtProInput.Height, blnCancel, False, True, UCase(txtProInput.Text) & "%")
'
'    If rsProvider.RecordCount = 0 Then
'        Exit Sub
'    End If
'    txtProInput.Text = rsProvider!名称
'End Sub

Private Sub cmdRestore_Click()
    Call GetValue(mlng药品ID)
End Sub

Private Sub cmdSave_Click()
    Dim strSql As String
    Dim i As Integer

    On Error GoTo errHandle
    With vsfUnit
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 1) = "" And .TextMatrix(i, 2) <> "" Then
                MsgBox "第" & i & "行厂家名称列为空了，请通过添加按钮添加值!", vbExclamation, gstrSysName
                Exit Sub
            End If
        Next

        For i = 1 To .Rows - 1
            strSql = strSql & .TextMatrix(i, 1) + "^" + .TextMatrix(i, 2) + "|"
        Next
        strSql = "Zl_药品生产商对照_Insert(" & lblMedi.Tag & ",'" & strSql & "')"

        zlDatabase.ExecuteProcedure strSql, "保存"
    End With
    lblSave.Visible = True
    mblnSave = True
    txtMedi.SetFocus
    zlControl.TxtSelAll txtMedi
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
    Dim rsTemp As ADODB.Recordset
    
    err = 0: On Error GoTo ErrHand
    
    gstrSql = "select I.ID,I.编码,I.名称,I.规格,I.产地,I.计算单位 as 单位" & _
            " from 收费项目目录 I" & _
            " where I.类别=[1] and I.ID=[2] " & _
            "       and (I.撤档时间 is null or I.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mstr类型, Val(mlng药品ID))
    
    With rsTemp
        If .BOF Or .EOF = 1 Then
            Me.lblMedi.Tag = 0: Me.txtMedi.Tag = "": Me.txtMedi.Text = Me.txtMedi.Tag
        Else
            Me.lblMedi.Tag = !ID
            Me.txtMedi.Tag = "[" & !编码 & "]" & !名称
            Me.txtMedi.Text = Me.txtMedi.Tag
            Me.lblSpec.Caption = "规格：" & IIf(IsNull(!规格), "", !规格) & _
                "   单位：" & IIf(IsNull(!单位), "", !单位)
        End If
    End With
    Me.txtMedi.SetFocus
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub InitVsf()
    With vsfUnit
        .Editable = flexEDKbdMouse
        .ColComboList(.ColIndex("厂家名称")) = ""
    End With
End Sub

Private Sub Form_Load()
    Me.Tag = mstr类型
    Me.lblMedi.Tag = mlng药品ID
    
    lblSave.Visible = False
    Me.lvwItems.ListItems.Clear
    With Me.lvwItems.ColumnHeaders
        .Clear
        .Add , "名称", "名称", 2000
        .Add , "编码", "编码", 1000
        .Add , "规格", "规格", 1200
        .Add , "产地", "生产商", 1200
        .Add , "单位", "单位", 600
    End With
    With Me.lvwItems
        .ColumnHeaders("编码").Position = 1
        .SortKey = .ColumnHeaders("编码").Index - 1
        .SortOrder = lvwAscending
    End With
    
    Call InitVsf
    Call GetValue(mlng药品ID)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstr原值 = ""
    mblnSave = False
End Sub

Private Sub lvwItems_DblClick()
    Dim intRow As Integer
    Dim intCol As Integer
    Dim strTemp As String
    
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    If mblnSave = False Then
        strTemp = ""
        With vsfUnit
            For intRow = 1 To .Rows - 1
                For intCol = 1 To .Cols - 1
                    strTemp = strTemp & .TextMatrix(intRow, intCol) & "|"
                Next
            Next
        End With
        If strTemp <> mstr原值 Then
            If MsgBox("当前内容被修改后未保存，你确定要继续吗？", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then lvwItems.Visible = False: Exit Sub
        End If
    End If
    
    With Me.lvwItems
        If Me.lblMedi.Tag <> Mid(.SelectedItem.Key, 2) Then
            Me.lblMedi.Tag = Mid(.SelectedItem.Key, 2)
            Me.txtMedi.Tag = "[" & .SelectedItem.SubItems(.ColumnHeaders("编码").Index - 1) & "]" & .SelectedItem.Text
            Me.txtMedi.Text = Me.txtMedi.Tag
            Me.lblSpec.Caption = "规格：" & .SelectedItem.SubItems(.ColumnHeaders("规格").Index - 1) & _
                        "   单位：" & .SelectedItem.SubItems(.ColumnHeaders("单位").Index - 1)
            mlng药品ID = Mid(.SelectedItem.Key, 2)
        End If
        Me.txtMedi.SetFocus
        Call zlCommFun.PressKey(vbKeyTab)
    End With
    Call GetValue(mlng药品ID)
    
    mblnSave = False
    lblSave.Visible = False
End Sub

Private Sub lvwItems_LostFocus()
    lvwItems.Visible = False
End Sub

Private Sub txtMedi_GotFocus()
    zlControl.TxtSelAll txtMedi
End Sub

Private Sub txtMedi_KeyPress(KeyAscii As Integer)
    Dim rsTemp As ADODB.Recordset
    Dim strTemp As String
    
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If KeyAscii <> vbKeyReturn Then Exit Sub
    strTemp = UCase(Trim(Me.txtMedi.Text))
    If strTemp = "" Then Me.lblMedi.Tag = 0: Me.txtMedi.Tag = "": Me.txtMedi.Text = "": Exit Sub
    
    If InStr(1, strTemp, "[") <> 0 And InStr(1, strTemp, "]") <> 0 Then strTemp = Mid(strTemp, 2, InStr(1, strTemp, "]") - 2)
    err = 0: On Error GoTo ErrHand
    
    gstrSql = "select distinct I.ID,I.编码,I.名称,I.规格,I.产地,I.计算单位 as 单位" & _
            " from 收费项目目录 I,收费项目别名 N" & _
            " where I.ID=N.收费细目ID and I.类别=[1] " & _
            "       and (I.撤档时间 is null or I.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))" & _
            "       and (I.编码 like [2] or N.名称 like [3] or N.简码 like [3])"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Me.Tag, strTemp & "%", gstrMatch & strTemp & "%")
    
    With rsTemp
        If .BOF Or .EOF = 1 Then
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
                        "   单位：" & IIf(IsNull(!单位), "", !单位)
                Else
                    Me.lblSpec.Caption = "   单位：" & IIf(IsNull(!单位), "", !单位)
                End If

            End If
            Call zlCommFun.PressKey(vbKeyTab)
            Call GetValue(!ID)
            Exit Sub
        End If
        Me.lvwItems.ListItems.Clear
        Do While Not .EOF
            Set mObjItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !名称)
            mObjItem.Icon = "ItemUse": mObjItem.SmallIcon = "ItemUse"
            mObjItem.SubItems(Me.lvwItems.ColumnHeaders("编码").Index - 1) = !编码
            mObjItem.SubItems(Me.lvwItems.ColumnHeaders("规格").Index - 1) = IIf(IsNull(!规格), "", !规格)
            mObjItem.SubItems(Me.lvwItems.ColumnHeaders("产地").Index - 1) = IIf(IsNull(!产地), "", !产地)
            mObjItem.SubItems(Me.lvwItems.ColumnHeaders("单位").Index - 1) = IIf(IsNull(!单位), "", !单位)
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

Private Sub vsfUnit_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsfUnit
        If NewCol = .ColIndex("批准文号") Then
            .FocusRect = flexFocusSolid
        Else
            .FocusRect = flexFocusLight
        End If
    End With
End Sub

Private Sub vsfUnit_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfUnit
        If .Cell(flexcpBackColor, Row, Col) = &H8000000F Then
            Cancel = True
        End If
    End With
End Sub

Private Sub GetValue(ByVal lngId As Long)
    '查询数据
    Dim strSql As String
    Dim i As Integer
    Dim j As Integer
    Dim intRow As Integer
    Dim intCol As Integer
    Dim rsRecord As ADODB.Recordset
    
    On Error GoTo errHandle
    mstr原值 = ""
    strSql = "select 厂家名称,批准文号 from 药品生产商对照 where 药品id=[1]"
    Set rsRecord = zlDatabase.OpenSQLRecord(strSql, "查询数据", lngId)
    
    vsfUnit.Cell(flexcpBackColor, 1, 0, vsfUnit.Rows - 1, 0) = &H8000000F
    With vsfUnit     '清空
        .Rows = 2
        For j = 0 To .Cols - 1
            .TextMatrix(1, j) = ""
        Next
    End With
    If rsRecord.EOF Then
        With vsfUnit
            For intRow = 1 To .Rows - 1
                For intCol = 1 To .Cols - 1
                    mstr原值 = mstr原值 & .TextMatrix(intRow, intCol) & "|"
                Next
            Next
        End With
        Exit Sub
    End If
    
    mstr原值 = ""
    vsfUnit.Rows = rsRecord.RecordCount + 1
    For i = 1 To rsRecord.RecordCount
        With vsfUnit
            .TextMatrix(i, 0) = i
            .TextMatrix(i, 1) = IIf(IsNull(rsRecord!厂家名称), "", rsRecord!厂家名称)
            .TextMatrix(i, 2) = IIf(IsNull(rsRecord!批准文号), "", rsRecord!批准文号)
        End With
        rsRecord.MoveNext
    Next
    vsfUnit.Cell(flexcpBackColor, 1, 0, vsfUnit.Rows - 1, 0) = &H8000000F
        
    With vsfUnit
        For intRow = 1 To .Rows - 1
            For intCol = 1 To .Cols - 1
                mstr原值 = mstr原值 & .TextMatrix(intRow, intCol) & "|"
            Next
        Next
    End With
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsfUnit_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    
    Dim vRect As RECT, blnCancel As Boolean
    Dim strPro As String
    Dim dblLeft As Double
    Dim dblTop As Double
    Dim rsProvider As ADODB.Recordset
    
    vRect = zlControl.GetControlRect(vsfUnit.hwnd) '获取位置
    dblLeft = vRect.Left + vsfUnit.CellLeft
    dblTop = vRect.Top + vsfUnit.CellTop + vsfUnit.CellHeight + 3200
    

    gstrSql = "select rownum as id, 名称 from 药品生产商"
    Set rsProvider = zlDatabase.ShowSQLSelect(Me, gstrSql, 0, "产地", False, "", "", False, False, _
    True, dblLeft, dblTop, vsfUnit.Height, blnCancel, False, True)

    If rsProvider Is Nothing Then
        Exit Sub
    End If
    With vsfUnit
        .TextMatrix(.Row, .Col) = rsProvider!名称
    End With
End Sub

Private Sub vsfUnit_DblClick()
    With vsfUnit
        .EditCell
        .EditSelStart = 0
        .EditSelLength = Len(.TextMatrix(.Row, .Col)) * 2
    End With
End Sub

Private Sub vsfUnit_EnterCell()
    With vsfUnit
        If .Col = 1 Then
            .ColComboList(.Col) = ""
        End If
    End With
End Sub

Private Sub vsfUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim vRect As RECT, blnCancel As Boolean
    Dim strPro As String
    Dim dblLeft As Double
    Dim dblTop As Double
    Dim rsProvider As ADODB.Recordset
    
    With vsfUnit
        If KeyCode = vbKeyReturn Then
            If .Col <> .Cols - 1 Then
                .Col = .Col + 1
            Else
                If .Row <> .Rows - 1 Then
                    .Row = .Row + 1
                    .Col = 1
                Else
                    If Trim(.TextMatrix(.Row, 1)) = "" Then KeyCode = 0: Exit Sub
                    .Rows = .Rows + 1
                    .Row = .Row + 1
                    .Col = 1
                    .TextMatrix(.Row, 0) = .Row
                    If .TextMatrix(1, 0) = "" Then
                        .TextMatrix(1, 0) = 1
                    End If
                    .Cell(flexcpBackColor, .Row, 0, .Row, 0) = &H8000000F
                End If
            End If
        ElseIf KeyCode = vbKeyDelete Then
            If .Rows <> 2 Then
                .RemoveItem .Row
            Else
                .TextMatrix(1, 1) = ""
                .TextMatrix(1, 2) = ""
            End If
        End If
    End With
End Sub

Private Sub vsfUnit_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    With vsfUnit
        If Col = 1 Then
            .ColComboList(Col) = "|..."
        End If
    End With
End Sub

Private Sub vsfUnit_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack Then
        With vsfUnit
            If Col = 2 And LenB(StrConv(.EditText, vbFromUnicode)) >= 40 Then
                KeyAscii = 0
            End If
        End With
    End If
End Sub

Private Sub vsfUnit_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    With vsfUnit
        If .Col = 1 Then
            .ColComboList(.Col) = "|..."
        Else
            .ColComboList(1) = ""
        End If
    End With
End Sub


Private Sub vsfUnit_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim vRect As RECT, blnCancel As Boolean
    Dim strPro As String
    Dim dblLeft As Double
    Dim dblTop As Double
    Dim rsProvider As ADODB.Recordset
    Dim i As Integer

    vRect = zlControl.GetControlRect(vsfUnit.hwnd) '获取位置
    dblLeft = vRect.Left + vsfUnit.CellLeft
    dblTop = vRect.Top + vsfUnit.CellTop + vsfUnit.CellHeight + 3200
    lblSave.Visible = False
    With vsfUnit
        If .Col = 1 And Trim(.EditText) <> "" Then
            gstrSql = "select rownum as id, 名称 from 药品生产商 where 名称 like [1] or 简码 like [1] or 编码 like [1]"
            Set rsProvider = zlDatabase.OpenSQLRecord(gstrSql, "生产商查询", UCase(.EditText) & "%")
            Set rsProvider = zlDatabase.ShowSQLSelect(Me, gstrSql, 0, "产地", False, "", "", False, False, _
            True, dblLeft, dblTop, vsfUnit.Height, blnCancel, False, True, UCase(Trim(.EditText)) & "%")

            If rsProvider Is Nothing Then
                MsgBox "无该生产商！", vbInformation, gstrSysName
                Cancel = True
                Exit Sub
            End If
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 1) = rsProvider!名称 And i <> Row Then
                    MsgBox "列表中已有该单位！", vbInformation, gstrSysName
                    Cancel = True
                    Exit Sub
                End If
            Next
            .EditText = rsProvider!名称
            .TextMatrix(.Row, .Col) = .EditText
        End If
    End With
End Sub
