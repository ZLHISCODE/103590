VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmMediUnit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "中标单位"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7200
   Icon            =   "frmMediUnit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdMedi 
      Caption         =   "…"
      Height          =   285
      Left            =   6840
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   818
      Width           =   285
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "保存(&S)"
      Height          =   350
      Left            =   4920
      TabIndex        =   4
      Top             =   4980
      Width           =   1095
   End
   Begin VB.CommandButton cmdRestore 
      Caption         =   "恢复(&R)"
      Height          =   350
      Left            =   1185
      TabIndex        =   6
      Top             =   4980
      Width           =   1095
   End
   Begin VB.CommandButton cmdDelAll 
      Caption         =   "全部删除(&A)"
      Height          =   350
      Left            =   2280
      TabIndex        =   7
      Top             =   4980
      Width           =   1245
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "删除(&D)"
      Height          =   350
      Left            =   90
      TabIndex        =   5
      Top             =   4980
      Width           =   1100
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "关闭(&C)"
      Height          =   350
      Left            =   6000
      TabIndex        =   8
      Top             =   4980
      Width           =   1100
   End
   Begin VB.TextBox txtMedi 
      Height          =   300
      Left            =   1260
      MaxLength       =   50
      TabIndex        =   1
      Top             =   810
      Width           =   5580
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   6105
      Top             =   5490
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
            Picture         =   "frmMediUnit.frx":1CFA
            Key             =   "ItemUse"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMediUnit.frx":2294
            Key             =   "ItemStop"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwItems 
      Height          =   2790
      Left            =   480
      TabIndex        =   11
      Top             =   5640
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
   Begin VSFlex8Ctl.VSFlexGrid vsfUnit 
      Height          =   3375
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   6975
      _cx             =   12303
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
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmMediUnit.frx":282E
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Msf中标单位选择 
      Height          =   2565
      Left            =   600
      TabIndex        =   12
      Top             =   1920
      Visible         =   0   'False
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   4524
      _Version        =   393216
      FixedCols       =   0
      GridColor       =   -2147483631
      GridColorFixed  =   8421504
      AllowBigSelection=   0   'False
      FocusRect       =   0
      FillStyle       =   1
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label lblSave 
      Caption         =   "保存成功！"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   3840
      TabIndex        =   13
      Top             =   5040
      Width           =   975
   End
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   120
      Picture         =   "frmMediUnit.frx":291F
      Top             =   0
      Width           =   480
   End
   Begin VB.Label lblSpec 
      AutoSize        =   -1  'True
      Caption         =   "规格：      生产商：       单位：瓶"
      Height          =   180
      Left            =   135
      TabIndex        =   10
      Top             =   1200
      Width           =   3150
   End
   Begin VB.Label lblnote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "    请选择药品后，指定其中标单位。招标药品入库时，其供应商必须属于中标单位"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   180
      Width           =   5685
   End
   Begin VB.Label lblMedi 
      AutoSize        =   -1  'True
      Caption         =   "药品规格(&M)"
      Height          =   180
      Left            =   120
      TabIndex        =   9
      Top             =   870
      Width           =   990
   End
End
Attribute VB_Name = "frmMediUnit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public lblTag As String
Public frmTag As String
Public strPrivs As String
Dim strTemp As String
Dim objItem As ListItem
Dim rsTemp As New ADODB.Recordset
Dim mblnStar As Boolean
Private mlngId As Long      '记录选中id
Private mstr原值 As String
Private mblnSave As Boolean     '记录是否保存了界面中修改的值

'记录状态类型
Private Enum mStates
    原始 = 0
    新增 = 1
    修改 = 2
    删除 = 3
End Enum

Private Const mcstIniColor = &H80000005
Private Const mcstUpdateColor = &HC2CBFE
Private Const mcstInsertColor = &HC2CBFE
Private Const mcstDelColor = &HDBDBDB
Private Sub vsf_ResetSerial()
    Dim i As Integer
    
    With vsfUnit
        For i = 1 To .Rows - 1
            .TextMatrix(i, 0) = i
        Next
    End With
End Sub

Private Sub cmdClose_Click()
    Dim intRow As Integer
    Dim intCol As Integer
    Dim strTemp As String

    With vsfUnit
        For intRow = 1 To .Rows - 1
            For intCol = 1 To .Cols - 1
                strTemp = strTemp & .TextMatrix(intRow, intCol) & "|"
            Next
        Next
    End With
    
    If strTemp <> mstr原值 Then
        If mblnSave = False Then
            If MsgBox("当前内容被修改后未保存，你确定要继续吗？", vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                mblnSave = False
                Unload Me
            End If
        Else
            Unload Me
        End If
    Else
        mblnSave = False
        Unload Me
    End If
End Sub

Private Sub cmdDel_Click()
    Dim i As Integer
    
    lblSave.Visible = False
    With vsfUnit
        If .Rows = 1 Then Exit Sub
        If Val(.TextMatrix(.Row, .ColIndex("单位ID"))) = 0 Then Exit Sub
        
        Select Case Val(.TextMatrix(.Row, .ColIndex("状态")))
            Case mStates.新增
                If .Rows - 1 = 1 Then
                    For i = 1 To .Cols - 1
                        .TextMatrix(1, i) = ""
                    Next
                Else
                    .RemoveItem .Row
                    vsf_ResetSerial
                End If
            Case mStates.原始
                .TextMatrix(.Row, .ColIndex("操作")) = mStates.删除
                .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = mcstDelColor
                cmdRestore.Enabled = True
                cmdDel.Enabled = False
        End Select
    End With
End Sub

Private Sub cmdDelAll_Click()
    Dim i As Integer
    
    lblSave.Visible = False
    With vsfUnit
        If .Rows = 1 Then Exit Sub
        .Redraw = flexRDNone
        For i = .Rows - 1 To 1 Step -1
            If Val(.TextMatrix(i, .ColIndex("单位ID"))) > 0 Then
                Select Case Val(.TextMatrix(i, .ColIndex("状态")))
                    Case mStates.新增
                        .RemoveItem i
                    Case mStates.原始
                        .TextMatrix(i, .ColIndex("操作")) = mStates.删除
                        .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = mcstDelColor
                End Select
            End If
        Next
        vsf_ResetSerial
        .Redraw = flexRDDirect
    End With
End Sub
Private Sub cmdMedi_Click()
    err = 0: On Error GoTo ErrHand
    
    gstrSql = "select I.ID,I.编码,I.名称,I.规格,I.产地,I.计算单位 as 单位" & _
            " from 收费项目目录 I" & _
            " where I.类别=[1] " & _
            "       and (I.撤档时间 is null or I.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, Me.Tag)
    
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
                        "   生产商：" & IIf(IsNull(!产地), "", !产地) & _
                        "   单位：" & IIf(IsNull(!单位), "", !单位)
                Else
                    Me.lblSpec.Caption = "生产商：" & IIf(IsNull(!产地), "", !产地) & "   单位：" & IIf(IsNull(!单位), "", !单位)
                End If
                Call ShowData
            End If
            Call OS.PressKey(vbKeyTab)
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

Private Sub cmdRestore_Click()
    lblSave.Visible = False
    
    With vsfUnit
        If .Rows = 1 Then Exit Sub
        If .Row = 0 Then Exit Sub
        If Val(.TextMatrix(.Row, .ColIndex("单位ID"))) = 0 Then Exit Sub
        
        .TextMatrix(.Row, .ColIndex("操作")) = .TextMatrix(.Row, .ColIndex("状态"))
        .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = mcstIniColor
        cmdRestore.Enabled = False
        cmdDel.Enabled = True
    End With
End Sub

Private Sub cmdSave_Click()
    Dim lngUnitId As Long
    Dim lngMediId As Long
    Dim str中标序号 As String
    Dim strDelDate As String
    Dim i As Integer

    On Error GoTo ErrHand
    
    mblnSave = True
    If vsfUnit.Rows = 1 Then Exit Sub
    lngMediId = Val(lblMedi.Tag)
    
    With vsfUnit
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("单位ID"))) > 0 Then
                lngUnitId = Val(.TextMatrix(i, .ColIndex("单位ID")))
                str中标序号 = .TextMatrix(i, .ColIndex("中标序号"))
                strDelDate = .TextMatrix(i, .ColIndex("建档时间"))
                
                gstrSql = ""
                Select Case Val(.TextMatrix(i, .ColIndex("操作")))
                    Case mStates.新增
                        gstrSql = "ZL_药品中标单位_INSERT(" & lngMediId & "," & lngUnitId & ", '" & str中标序号 & "')"
                    Case mStates.修改
                        gstrSql = "Zl_药品中标单位_Update(" & lngMediId & "," & lngUnitId & ",to_date('" & strDelDate & "','YYYY-MM-DD HH24:MI:SS') , '" & str中标序号 & "')"
                    Case mStates.删除
                        gstrSql = "ZL_药品中标单位_DELETE(" & lngMediId & "," & lngUnitId & ",to_date('" & strDelDate & "','YYYY-MM-DD HH24:MI:SS'))"
                End Select
                
                If gstrSql <> "" Then Call zldatabase.ExecuteProcedure(gstrSql, Me.Caption)
            
                '同步物流平台药品信息
                If Not gobjLogisticPlatform Is Nothing Then
                    If Val(.TextMatrix(i, .ColIndex("操作"))) = mStates.删除 Then
                        gobjLogisticPlatform.ClearDrugInfo lngMediId, lngUnitId
                    End If
                    gobjLogisticPlatform.UploadDrugInfo Me, gcnOracle, lngMediId
                End If
            
            End If
        Next
    End With
    
    lblSave.Visible = True
    Call ShowData
    txtMedi.SetFocus
'    txtProInput.SetFocus
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
    If Me.cmdClose.Tag = "查阅" Then
'        Me.cmdSave.Visible = False
    End If
    
    err = 0: On Error GoTo ErrHand
    
    gstrSql = "select I.ID,I.编码,I.名称,I.规格,I.产地,I.计算单位 as 单位" & _
            " from 收费项目目录 I" & _
            " where I.类别=[1] and I.ID=[2] " & _
            "       and (I.撤档时间 is null or I.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, Me.Tag, Val(Me.lblMedi.Tag))
    
    With rsTemp
        If .BOF Or .EOF = 1 Then
            Me.lblMedi.Tag = 0: Me.txtMedi.Tag = "": Me.txtMedi.Text = Me.txtMedi.Tag
        Else
            Me.lblMedi.Tag = !ID
            Me.txtMedi.Tag = "[" & !编码 & "]" & !名称
            Me.txtMedi.Text = Me.txtMedi.Tag
            Me.lblSpec.Caption = "规格：" & IIf(IsNull(!规格), "", !规格) & _
                "   生产商：" & IIf(IsNull(!产地), "", !产地) & _
                "   单位：" & IIf(IsNull(!单位), "", !单位)
            Call ShowData
        End If
    End With
'    Me.txtProInput.SetFocus
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyEscape
        If Msf中标单位选择.Visible Then
            Msf中标单位选择.Visible = False
            Exit Sub
        End If
        If lvwItems.Visible Then
            lvwItems.Visible = False: txtMedi.SetFocus: Exit Sub
        End If
        Call cmdClose_Click
    Case Else
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    '中草药不允许设置中标单位
    On Error Resume Next
    
    Me.Tag = frmTag
    Me.lblMedi.Tag = lblTag
    lblSave.Visible = False
    
'    If Me.Tag = "7" Then
'        MsgBox "中草药不允许设置中标单位！", vbInformation, gstrSysName
'        Unload Me
'        Exit Sub
'    End If
    If InStr(1, strPrivs, "中标单位") = 0 Then
        MsgBox "你不具备管理中标单位的权限！", vbInformation, gstrSysName
        Unload Me
        Exit Sub
    End If
    
    Me.lvwItems.ListItems.Clear
    With Me.lvwItems.ColumnHeaders
        .Clear
        .Add , "名称", "名称", 2000
        .Add , "编码", "编码", 1000
        .Add , "规格", "规格", 1200
        .Add , "产地", IIf(Me.Tag = "7", "产地", "厂牌"), 1200
        .Add , "单位", "单位", 600
    End With
    With Me.lvwItems
        .ColumnHeaders("编码").Position = 1
        .SortKey = .ColumnHeaders("编码").Index - 1
        .SortOrder = lvwAscending
    End With
    With vsfUnit
        .ColComboList(.ColIndex("单位")) = "|..."
        .Editable = flexEDKbdMouse
    End With
    
    Call ShowData
End Sub

Private Sub vsfUnit_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsRecord As ADODB.Recordset
    Dim vRect As RECT, blnCancel As Boolean
    Dim dblLeft As Double
    Dim dblTop As Double
    
    vRect = zlControl.GetControlRect(vsfUnit.hWnd) '获取位置
    dblLeft = vRect.Left + vsfUnit.CellLeft
    dblTop = vRect.Top + vsfUnit.CellTop + vsfUnit.CellHeight + 3200
    With vsfUnit
        If Col = .ColIndex("单位") Then
            gstrSql = "Select ID,编码,名称,简码 From 供应商 Where 末级=1 And (instr(类型,1,1)=1 Or Nvl(末级,0)=0) And (撤档时间 is null or 撤档时间=to_date('3000-01-01','YYYY-MM-DD')) Order By 编码 "
            
            Set rsRecord = zldatabase.ShowSQLSelect(Me, gstrSql, 0, "供应商", False, "", "", False, False, _
            True, dblLeft, dblTop, .Height, blnCancel, False, True)

            If rsRecord Is Nothing Then
                Exit Sub
            Else
                mlngId = 0
                mlngId = rsRecord!ID
                If CheckDub = False Then
                    .TextMatrix(Row, 0) = Row
                    .TextMatrix(Row, .ColIndex("单位")) = "[" & rsRecord!编码 & "]" & rsRecord!名称
                    .TextMatrix(Row, .ColIndex("状态")) = mStates.新增
                    .TextMatrix(Row, .ColIndex("操作")) = mStates.新增
                    .TextMatrix(Row, .ColIndex("单位ID")) = rsRecord!ID
    '                .Cell(flexcpBackColor, Row, 0, Row, .Cols - 1) = vbWhite  'mcstInsertColor
                    .Cell(flexcpBackColor, Row, .ColIndex("建档时间"), Row, .ColIndex("建档时间")) = mcstDelColor
                    .Col = .ColIndex("中标序号")
                    lblSave.Visible = False
                Else
                    MsgBox "已经有该中标单位！", vbInformation, gstrSysName
                End If
            End If
        End If
    End With
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
    err = 0: On Error GoTo ErrHand
    
    gstrSql = "select distinct I.ID,I.编码,I.名称,I.规格,I.产地,I.计算单位 as 单位" & _
            " from 收费项目目录 I,收费项目别名 N" & _
            " where I.ID=N.收费细目ID and I.类别=[1] " & _
            "       and (I.撤档时间 is null or I.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))" & _
            "       and (I.编码 like [2] or N.名称 like [3] or N.简码 like [3])"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, Me.Tag, strTemp & "%", gstrMatch & strTemp & "%")
    
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
                        "   生产商：" & IIf(IsNull(!产地), "", !产地) & _
                        "   单位：" & IIf(IsNull(!单位), "", !单位)
                Else
                    Me.lblSpec.Caption = "生产商：" & IIf(IsNull(!产地), "", !产地) & "   单位：" & IIf(IsNull(!单位), "", !单位)
                End If
                Call ShowData
            End If
            Call OS.PressKey(vbKeyTab)
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

Private Sub ShowData()
    Dim intRow As Integer
    Dim intCol As Integer
    Dim i As Integer
    
    On Error GoTo errHandle
    mblnStar = True
    '显示已初始化的中标单位
    vsfUnit.TextMatrix(1, 0) = "1"
    
    gstrSql = "Select C.ID,'['||C.编码||']'||C.名称 单位,B.建档时间,B.中标序号 From 药品规格 A,药品中标单位 B,供应商 C" & _
            " Where A.药品ID=B.药品ID And instr(C.类型,1,1)=1 And B.单位ID=C.ID And A.药品ID=[1] " & _
            " And (B.撤档时间 is null or B.撤档时间=to_date('3000-01-01','YYYY-MM-DD')) " & _
            " Order by B.建档时间"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(lblMedi.Tag))
    
    With rsTemp
        vsfUnit.Rows = 1
        vsfUnit.Rows = IIf(.RecordCount > 0, .RecordCount + 1, 2)
        Do While Not .EOF
            vsfUnit.TextMatrix(.AbsolutePosition, 0) = .AbsolutePosition
            vsfUnit.TextMatrix(.AbsolutePosition, vsfUnit.ColIndex("单位")) = !单位
            vsfUnit.TextMatrix(.AbsolutePosition, vsfUnit.ColIndex("中标序号")) = IIf(IsNull(!中标序号), "", !中标序号)
            vsfUnit.TextMatrix(.AbsolutePosition, vsfUnit.ColIndex("建档时间")) = Format(!建档时间, "YYYY-MM-DD HH:MM:SS")
            vsfUnit.TextMatrix(.AbsolutePosition, vsfUnit.ColIndex("状态")) = 0
            vsfUnit.TextMatrix(.AbsolutePosition, vsfUnit.ColIndex("操作")) = 0
            vsfUnit.TextMatrix(.AbsolutePosition, vsfUnit.ColIndex("单位ID")) = !ID
            .MoveNext
        Loop
    End With
    
    mstr原值 = ""
    With vsfUnit
        .Cell(flexcpBackColor, 0, .ColIndex("建档时间"), .Rows - 1, .ColIndex("建档时间")) = mcstDelColor
        For intRow = 1 To .Rows - 1
            For intCol = 1 To .Cols - 1
                mstr原值 = mstr原值 & .TextMatrix(intRow, intCol) & "|"
            Next
        Next
    End With
    mblnStar = False
        
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
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
    Dim intRow As Integer
    Dim intCol As Integer
    Dim strTemp As String
    
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    With vsfUnit
        For intRow = 1 To .Rows - 1
            For intCol = 1 To .Cols - 1
                strTemp = strTemp & .TextMatrix(intRow, intCol) & "|"
            Next
        Next
    End With
    
    If strTemp <> mstr原值 Then
        If mblnSave = False Then
            If MsgBox("当前内容被修改后未保存，你确定要继续吗？", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                lvwItems.Visible = False
                Exit Sub
            End If
        End If
    End If
    
    With Me.lvwItems
        If Me.lblMedi.Tag <> Mid(.SelectedItem.Key, 2) Then
            Me.lblMedi.Tag = Mid(.SelectedItem.Key, 2)
            Me.txtMedi.Tag = "[" & .SelectedItem.SubItems(.ColumnHeaders("编码").Index - 1) & "]" & .SelectedItem.Text
            Me.txtMedi.Text = Me.txtMedi.Tag
            lblSpec.Caption = "规格：" & lvwItems.SelectedItem.SubItems(.ColumnHeaders("规格").Index - 1) & _
                                "      生产商：" & lvwItems.SelectedItem.SubItems(3) & _
                                "     单位：" & lvwItems.SelectedItem.SubItems(.ColumnHeaders("单位").Index - 1)
            Call ShowData
        End If
        Me.txtMedi.SetFocus
        Call OS.PressKey(vbKeyTab)
    End With
    lblSave.Visible = False
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

Private Sub msf中标单位选择_LostFocus()
    With Msf中标单位选择
        .ZOrder 1
        .Visible = False
    End With
End Sub

Private Sub vsfUnit_EnterCell()
    With vsfUnit
        If .Rows = 1 Then Exit Sub
        If .Row = 0 Then Exit Sub
'        If Val(.TextMatrix(.Row, .ColIndex("单位ID"))) = 0 Then Exit Sub
        
        Select Case Val(.TextMatrix(.Row, .ColIndex("操作")))
            Case mStates.原始, mStates.新增
                cmdDel.Enabled = True
                cmdRestore.Enabled = False
            Case mStates.删除
                cmdDel.Enabled = False
                cmdRestore.Enabled = True
        End Select
        If .Cell(flexcpBackColor, .Row, .Col, .Row, .Col) = mcstDelColor Then
            .Editable = flexEDNone
        Else
            .Editable = flexEDKbdMouse
        End If
        If .Col = .ColIndex("单位") Then
            .ColComboList(.ColIndex("单位")) = ""
        End If
    End With
End Sub

Private Sub vsfUnit_GotFocus()
    lvwItems.Visible = False
End Sub

Private Sub vsfUnit_KeyDown(KeyCode As Integer, Shift As Integer)
    With vsfUnit
        If KeyCode = vbKeyReturn Then
            If .Col <> .ColIndex("中标序号") Then
                .Col = .Col + 1
            ElseIf .Row <> .Rows - 1 And .Col = .ColIndex("中标序号") Then
                .Row = .Row + 1
                .Col = .ColIndex("单位")
            ElseIf .Row = .Rows - 1 And .TextMatrix(.Row, 1) <> "" Then
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 0) = .Rows - 1
                .Row = .Rows - 1
                .Col = .ColIndex("单位")
            End If
        ElseIf KeyCode = vbKeyDelete Then
            Call cmdDel_Click
        End If
        If .Cell(flexcpBackColor, .Row, .Col, .Row, .Col) = mcstDelColor Then
            KeyCode = 0
        End If
    End With
End Sub

Private Sub vsfUnit_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    If Col = vsfUnit.ColIndex("单位") Then
        vsfUnit.ColComboList(Col) = "|..."
    End If
End Sub

Private Sub vsfUnit_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    With vsfUnit
        If .Col = .ColIndex("单位") Then
            .ColComboList(.ColIndex("单位")) = "|..."
        Else
            .ColComboList(.ColIndex("单位")) = ""
        End If
    End With
End Sub

Private Sub vsfUnit_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    With vsfUnit
        .EditSelStart = 0
        .EditSelLength = zlCommFun.ActualLen(.EditText)
    End With
End Sub

Private Sub vsfUnit_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfUnit
'        If Col <> .ColIndex("中标序号") Or Val(.TextMatrix(Row, .ColIndex("操作"))) = mStates.删除 Then Cancel = True
        If Col = .ColIndex("中标序号") Then
            .EditMaxLength = 50
        Else
            .EditMaxLength = 50
        End If
    End With
End Sub

Private Sub vsfUnit_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rsRecord As ADODB.Recordset
    Dim vRect As RECT, blnCancel As Boolean
    Dim dblLeft As Double
    Dim dblTop As Double
    Dim intRow As Integer
    Dim intCol As Integer
    
    vRect = zlControl.GetControlRect(vsfUnit.hWnd) '获取位置
    dblLeft = vRect.Left + vsfUnit.CellLeft
    dblTop = vRect.Top + vsfUnit.CellTop + vsfUnit.CellHeight + 3200
    lblSave.Visible = False
    With vsfUnit
        If .EditText = "" Then Exit Sub
        If .Rows - 1 >= 1 Then
        If Col = .ColIndex("单位") And InStr(1, .EditText, "[") = 0 Then
            gstrSql = "Select ID,编码,名称,简码 From 供应商 Where 末级=1 And (instr(类型,1,1)=1 Or Nvl(末级,0)=0) And (撤档时间 is null or 撤档时间=to_date('3000-01-01','YYYY-MM-DD')) and (编码 like [1] or 简码 like[1] or 名称 like [1]) Order By 编码 "
            
            Set rsRecord = zldatabase.ShowSQLSelect(Me, gstrSql, 0, "供应商", False, "", "", False, False, _
            True, dblLeft, dblTop, .Height, blnCancel, False, True, UCase(.EditText) & "%")

            If Not rsRecord Is Nothing Then
                mlngId = 0
                mlngId = rsRecord!ID
                If CheckDub = False Then
                    .TextMatrix(Row, 0) = Row
                    .EditText = "[" & rsRecord!编码 & "]" & rsRecord!名称
                    .TextMatrix(Row, .ColIndex("单位")) = .EditText
                    .TextMatrix(Row, .ColIndex("状态")) = mStates.新增
                    .TextMatrix(Row, .ColIndex("操作")) = mStates.新增
                    .TextMatrix(Row, .ColIndex("单位ID")) = rsRecord!ID
                    .Cell(flexcpBackColor, Row, .ColIndex("建档时间"), Row, .ColIndex("建档时间")) = mcstDelColor
                    .Col = .ColIndex("中标序号")
                Else
                    MsgBox "已经有该中标单位！", vbInformation, gstrSysName
                    Cancel = True
                End If
            Else
                MsgBox "无该中标单位！", vbInformation, gstrSysName
                Cancel = True
            End If
        End If
        End If
    End With
End Sub

Private Function CheckDub() As Boolean
    '检查是否存在该中标单位
    Dim i As Integer
    
    With vsfUnit
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("单位id")) <> "" Then
                If .TextMatrix(i, .ColIndex("单位id")) = mlngId Then
                    CheckDub = True
                    Exit Function
                End If
            Else
                CheckDub = False
            End If
        Next
    End With
    CheckDub = False
End Function
