VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmStuffUnitMgr 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "中标单位"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6720
   Icon            =   "frmStuffUnitMgr.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   6720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.OptionButton optApply 
      Caption         =   "应用于所有卫生材料"
      Height          =   255
      Index           =   3
      Left            =   3360
      TabIndex        =   17
      Top             =   5040
      Width           =   2055
   End
   Begin VB.OptionButton optApply 
      Caption         =   "应用于此分类下所有规格"
      Height          =   180
      Index           =   2
      Left            =   240
      TabIndex        =   16
      Top             =   5077
      Width           =   2535
   End
   Begin VB.OptionButton optApply 
      Caption         =   "应用于本品种所有规格"
      Height          =   255
      Index           =   1
      Left            =   3360
      TabIndex        =   15
      Top             =   4680
      Width           =   2295
   End
   Begin VB.OptionButton optApply 
      Caption         =   "应用于本规格"
      Height          =   180
      Index           =   0
      Left            =   240
      TabIndex        =   14
      Top             =   4717
      Value           =   -1  'True
      Width           =   1575
   End
   Begin VB.CommandButton cmdStuff 
      Caption         =   "…"
      Height          =   285
      Left            =   6240
      TabIndex        =   13
      TabStop         =   0   'False
      Tag             =   "分类"
      ToolTipText     =   "按*打开选择器"
      Top             =   818
      Width           =   285
   End
   Begin VB.CheckBox chk招标材料 
      Caption         =   "招标卫材(&U)"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   1425
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "关闭(&X)"
      Height          =   350
      Left            =   5430
      TabIndex        =   7
      Top             =   5415
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   210
      Picture         =   "frmStuffUnitMgr.frx":1CFA
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   5415
      Width           =   1100
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "保存(&S)"
      Height          =   350
      Left            =   4320
      TabIndex        =   6
      Top             =   5415
      Width           =   1100
   End
   Begin VB.TextBox txtStuff 
      Height          =   300
      Left            =   1260
      MaxLength       =   50
      TabIndex        =   2
      Top             =   810
      Width           =   4980
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "全部清除(&C)"
      Height          =   350
      Left            =   1410
      Picture         =   "frmStuffUnitMgr.frx":1E44
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   5415
      Width           =   1290
   End
   Begin VB.CommandButton cmdRestore 
      Caption         =   "恢复(&R)"
      Height          =   350
      Left            =   2700
      Picture         =   "frmStuffUnitMgr.frx":1F8E
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   5415
      Width           =   1290
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   5400
      Top             =   6120
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
            Picture         =   "frmStuffUnitMgr.frx":20D8
            Key             =   "ItemUse"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffUnitMgr.frx":2672
            Key             =   "ItemStop"
         EndProperty
      EndProperty
   End
   Begin ZL9BillEdit.BillEdit msfUnit 
      Height          =   2775
      Left            =   210
      TabIndex        =   5
      Top             =   1695
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
   Begin MSComctlLib.ListView lvwItems 
      Height          =   2790
      Left            =   0
      TabIndex        =   11
      Top             =   6360
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Msf中标单位选择 
      Height          =   2565
      Left            =   0
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
   Begin VB.Image imgNote 
      Height          =   480
      Left            =   240
      Picture         =   "frmStuffUnitMgr.frx":2C0C
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblSpec 
      AutoSize        =   -1  'True
      Caption         =   "规格：      厂牌：       单位：瓶"
      Height          =   180
      Left            =   1260
      TabIndex        =   3
      Top             =   1200
      Width           =   2970
   End
   Begin VB.Label lblnote 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "    请选择卫材后，指定其中标单位。招标卫材入库时，其供应商必须属于中标单位"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   180
      Width           =   5685
   End
   Begin VB.Label lblStuff 
      AutoSize        =   -1  'True
      Caption         =   "指定卫材(&M)"
      Height          =   180
      Left            =   240
      TabIndex        =   1
      Top             =   870
      Width           =   990
   End
End
Attribute VB_Name = "frmStuffUnitMgr"
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

Private Sub chk招标材料_Click()
    msfUnit.Active = (chk招标材料.Value = 1)
    If chk招标材料.Value = 0 Then
        Call cmdClear_Click
    Else
        Call cmdRestore_Click
    End If
End Sub

Private Sub chk招标材料_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey (vbKeyTab)
End Sub

Private Sub cmdClear_Click()
    msfUnit.ClearBill
    msfUnit.TextMatrix(1, 0) = "1"
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int(glngSys / 100))
End Sub

Private Sub cmdStuff_Click()

    err = 0: On Error GoTo ErrHand
    
    gstrSQL = "select I.ID,I.编码,I.名称,I.规格,I.产地,I.计算单位 as 单位" & _
            " from 收费项目目录 I" & _
            " where I.类别='4'" & _
            "       and (I.撤档时间 is null or I.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))"
    
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
    
    
    With rsTemp
        If .BOF Or .EOF = 1 Then
            MsgBox "尚未建立卫生材料档案信息！", vbExclamation, gstrSysName
            Me.lblStuff.Tag = 0: Me.txtStuff.Tag = "": Me.txtStuff.Text = Me.txtStuff.Tag: Me.txtStuff.SetFocus: Exit Sub
        End If
        If .RecordCount = 1 Then
            If Me.lblStuff.Tag <> !Id Then
                Me.lblStuff.Tag = !Id
                Me.txtStuff.Tag = "[" & !编码 & "]" & !名称
                Me.txtStuff.Text = Me.txtStuff.Tag
                If Me.Tag <> "7" Then
                    Me.lblSpec.Caption = "规格：" & IIf(IsNull(!规格), "", !规格) & _
                        "   厂牌：" & IIf(IsNull(!产地), "", !产地) & _
                        "   单位：" & IIf(IsNull(!单位), "", !单位)
                Else
                    Me.lblSpec.Caption = "产地：" & IIf(IsNull(!产地), "", !产地) & "   单位：" & IIf(IsNull(!单位), "", !单位)
                End If
                Call ShowData
            End If
            Call OS.PressKey(vbKeyTab)
            Exit Sub
        End If
        Me.lvwItems.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwItems.ListItems.Add(, "_" & !Id, !名称)
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
        .Tag = Me.txtStuff.Name
        .Left = Me.txtStuff.Left
        .Top = Me.txtStuff.Top + Me.txtStuff.Height
        .ZOrder 0: .Visible = True
        .SetFocus
    End With
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdRestore_Click()
    Call ShowData
End Sub

Private Sub CmdSave_Click()
    Dim lngRow As Long
    Dim str单位 As String
    Dim intApply As Integer
    Dim i As Integer
    
    On Error GoTo ErrHand
    
    If Val(Me.lblStuff.Tag) = 0 Then
        MsgBox "还未选择卫生材料，不能保存！", vbInformation, gstrSysName
        txtStuff.SetFocus
        Exit Sub
    End If
    
    For lngRow = 1 To msfUnit.Rows - 1
        If Val(msfUnit.TextMatrix(lngRow, 3)) > 1000000 Then
            MsgBox "第" & lngRow & "行成本价超过最大值1000000，不能保存！", vbInformation, gstrSysName
            Exit Sub
        End If
    Next
    
    If optApply(0).Value = False Then
        For i = 0 To optApply.UBound
            If optApply(i).Value = True Then
                If MsgBox("该卫材中标单位应用范围为“" & optApply(i).Caption & "”是否继续？", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                Else
                    Exit For
                End If
            End If
        Next
    End If
    
    'str单位格式：单位id,中标序号|单位id,中标序号....
    With msfUnit
        For lngRow = 1 To .Rows - 1
            If Val(.RowData(lngRow)) <> 0 Then
                str单位 = IIf(str单位 = "", "", str单位 & "|") & Val(.RowData(lngRow)) & "," & .TextMatrix(lngRow, 2) & "," & .TextMatrix(lngRow, 3)
            End If
        Next
    End With
    
    For i = 0 To optApply.UBound
        If optApply(i).Value = True Then
            intApply = i
            Exit For
        End If
    Next
    
    gstrSQL = "ZL_材料中标单位_INSERT(" & Val(lblStuff.Tag) & ",'" & str单位 & "'," & intApply & ")"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    
    Unload Me
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
    If Me.cmdClose.Tag = "查阅" Then
        Me.msfUnit.Active = False
        Me.cmdSave.Visible = False
        Me.cmdClear.Visible = False
        Me.cmdRestore.Visible = False
        Me.cmdStuff.Enabled = False
        Me.txtStuff.Enabled = False
        Me.chk招标材料.Enabled = False
        
        
    End If
    
    err = 0: On Error GoTo ErrHand
    gstrSQL = "select I.ID,I.编码,I.名称,I.规格,I.产地,I.计算单位 as 单位" & _
            " from 收费项目目录 I" & _
            " where I.类别='4' and I.ID=[1]" & _
            "       and (I.撤档时间 is null or I.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(Me.lblStuff.Tag))
    If rsTemp.State <> 1 Then
            Me.lblStuff.Tag = 0: Me.txtStuff.Tag = "": Me.txtStuff.Text = Me.txtStuff.Tag
    End If
    With rsTemp
        If .BOF Or .EOF = 1 Then
            Me.lblStuff.Tag = 0: Me.txtStuff.Tag = "": Me.txtStuff.Text = Me.txtStuff.Tag
        Else
            Me.lblStuff.Tag = !Id
            Me.txtStuff.Tag = "[" & !编码 & "]" & !名称
            Me.txtStuff.Text = Me.txtStuff.Tag
            Me.lblSpec.Caption = "规格：" & IIf(IsNull(!规格), "", !规格) & _
                "   厂牌：" & IIf(IsNull(!产地), "", !产地) & _
                "   单位：" & IIf(IsNull(!单位), "", !单位)
            Call ShowData
        End If
    End With
    If Me.txtStuff.Enabled Then Me.txtStuff.SetFocus
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
            msfUnit.TxtSetFocus
            Exit Sub
        Else
            cmdClose_Click
        End If
    Case Else
    End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    On Error Resume Next
    
    Me.Tag = frmTag
    Me.lblStuff.Tag = lblTag
    
    With msfUnit
        .Rows = 2
        .Cols = 4
        
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "单位名称"
        .TextMatrix(0, 2) = "中标序号"
        .TextMatrix(0, 3) = "成本价"
        .TextMatrix(1, 0) = "1"
        .ColData(0) = 5
        .ColData(1) = 1
        .ColData(2) = 4
        .ColData(3) = 4
        .ColWidth(0) = 300
        .ColWidth(1) = 3500
        .ColWidth(2) = 1000
        .ColWidth(3) = 1000
        
        .PrimaryCol = 1
        .LocateCol = 1
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
    
    Call ShowData
End Sub

Private Sub msfUnit_EnterCell(Row As Long, Col As Long)
    With msfUnit
        If Col = 1 Then
            .TextMask = ""
        ElseIf Col = 2 Then
            .TxtCheck = True
            .TextMask = "1234567890"
            .MaxLength = 50
        ElseIf Col = 3 Then
            .TxtCheck = True
            .TextMask = "1234567890."
            .MaxLength = 50
        End If
    End With
End Sub

Private Sub msfUnit_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strFind As String
    Dim rsTemp As New ADODB.Recordset
    Dim strKey As String
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    On Error GoTo ErrHandle
    With msfUnit
        If .Col <> 1 Then Exit Sub
        If .TxtVisible = False Then Exit Sub
        If .Text = "" Then Exit Sub
        
        strKey = GetMatchingSting(UCase(.Text))
        strFind = " And (编码 Like [1]" & _
                    " Or upper(名称) Like [1]" & _
                    " Or 简码 Like [1])"
    End With
    
    gstrSQL = " Select ID,编码,名称,简码 From 供应商 " & _
             " Where 末级=1 And (substr(类型,5,1)=1 Or Nvl(末级,0)=0) " & strFind & " Order By 编码 "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strKey)
    With rsTemp
        If .EOF Then
            MsgBox "没有找到匹配的卫材供应商，请重新输入！", vbInformation, gstrSysName
            Cancel = True
            msfUnit.TxtSetFocus
            Exit Sub
        End If
        
        With Msf中标单位选择
            .Clear
            Set .DataSource = rsTemp
            .ColWidth(0) = 0
            .ColWidth(1) = 800
            .ColWidth(2) = 3000
            .ColWidth(3) = 800
            .Top = msfUnit.Top + msfUnit.CellTop + msfUnit.MsfObj.CellHeight
            If .Top + .Height > Me.Height Then .Top = msfUnit.Top + msfUnit.CellTop - .Height
            .Visible = True
            .ZOrder 0
            
            .Row = 1
            .ColSel = .Cols - 1
            .SetFocus
        End With
    End With
    Cancel = True
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub optApply_Click(Index As Integer)
    Dim i As Integer
    
    For i = 1 To optApply.UBound
        If i = Index Then
            optApply(i).FontBold = True
        Else
            optApply(i).FontBold = False
        End If
    Next
End Sub

Private Sub txtStuff_GotFocus()
    Me.txtStuff.SelStart = 0: Me.txtStuff.SelLength = 100
End Sub

Private Sub txtStuff_KeyPress(KeyAscii As Integer)
    If InStr(" ~!@#$%^&*()_+|=-`;'"":/.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If KeyAscii <> vbKeyReturn Then Exit Sub
    strTemp = UCase(Trim(Me.txtStuff.Text))
    If strTemp = "" Then Me.lblStuff.Tag = 0: Me.txtStuff.Tag = "": Me.txtStuff.Text = "": Exit Sub
    
    If InStr(1, strTemp, "[") <> 0 And InStr(1, strTemp, "]") <> 0 Then strTemp = Mid(strTemp, 2, InStr(1, strTemp, "]") - 2)
    
    err = 0: On Error GoTo ErrHand
    strTemp = GetMatchingSting(strTemp)
    
    gstrSQL = "select distinct I.ID,I.编码,I.名称,I.规格,I.产地,I.计算单位 as 单位" & _
            " from 收费项目目录 I,收费项目别名 N" & _
            " where I.ID=N.收费细目ID and I.类别='4'" & _
            "       and (I.撤档时间 is null or I.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))" & _
            "       and (I.编码 like [1] or N.名称 like [1] or N.简码 like [1])"
            
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strTemp)
    With rsTemp
        If .BOF Or .EOF = 1 Then
            MsgBox "未找到指定规格的卫生材料，请重新指定！", vbExclamation, gstrSysName
            Me.lblStuff.Tag = 0: Me.txtStuff.Tag = "": Me.txtStuff.Text = Me.txtStuff.Tag: Me.txtStuff.SetFocus: Exit Sub
        End If
        If .RecordCount = 1 Then
            If Me.lblStuff.Tag <> !Id Then
                Me.lblStuff.Tag = !Id
                Me.txtStuff.Tag = "[" & !编码 & "]" & !名称
                Me.txtStuff.Text = Me.txtStuff.Tag
                If Me.Tag <> "7" Then
                    Me.lblSpec.Caption = "规格：" & IIf(IsNull(!规格), "", !规格) & _
                        "   厂牌：" & IIf(IsNull(!产地), "", !产地) & _
                        "   单位：" & IIf(IsNull(!单位), "", !单位)
                Else
                    Me.lblSpec.Caption = "产地：" & IIf(IsNull(!产地), "", !产地) & "   单位：" & IIf(IsNull(!单位), "", !单位)
                End If
                Call ShowData
            End If
            Call OS.PressKey(vbKeyTab)
            Exit Sub
        End If
        Me.lvwItems.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwItems.ListItems.Add(, "_" & !Id, !名称)
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
        .Tag = Me.txtStuff.Name
        .Left = Me.txtStuff.Left
        .Top = Me.txtStuff.Top + Me.txtStuff.Height
        .ZOrder 0: .Visible = True
        .SetFocus
    End With
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtStuff_LostFocus()
    Me.txtStuff.Text = Me.txtStuff.Tag
End Sub

Private Sub ShowData()
    '显示已初始化的中标单位
    msfUnit.ClearBill
    msfUnit.TextMatrix(1, 0) = "1"
    
    On Error GoTo ErrHandle
    gstrSQL = "Select C.ID,'['||C.编码||']'||C.名称 单位,B.成本价,中标序号 From 材料特性 A,材料中标单位 B,供应商 C" & _
            " Where A.材料ID=B.材料ID And substr(C.类型,5,1)=1 And B.单位ID=C.ID And A.材料ID=[1]"
            
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(lblStuff.Tag))
    With rsTemp
        If .RecordCount <> 0 Then chk招标材料.Value = 1
        Do While Not .EOF
            msfUnit.TextMatrix(.AbsolutePosition, 0) = .AbsolutePosition
            msfUnit.TextMatrix(.AbsolutePosition, 1) = IIf(IsNull(!单位), "", !单位)
            msfUnit.TextMatrix(.AbsolutePosition, 2) = IIf(IsNull(!中标序号), "", !中标序号)
            msfUnit.TextMatrix(.AbsolutePosition, 3) = IIf(IsNull(!成本价), "", !成本价)
            msfUnit.RowData(.AbsolutePosition) = !Id
            If msfUnit.Rows - 1 >= .AbsolutePosition Then msfUnit.Rows = msfUnit.Rows + 1
            .MoveNext
        Loop
        If msfUnit.RowData(msfUnit.Rows - 1) = 0 And msfUnit.Rows > 2 Then msfUnit.Rows = msfUnit.Rows - 1
    End With
    Exit Sub
ErrHandle:
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
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    With Me.lvwItems
        If Me.lblStuff.Tag <> Mid(.SelectedItem.Key, 2) Then
            Me.lblStuff.Tag = Mid(.SelectedItem.Key, 2)
            Me.txtStuff.Tag = "[" & .SelectedItem.SubItems(.ColumnHeaders("编码").Index - 1) & "]" & .SelectedItem.Text
            Me.txtStuff.Text = Me.txtStuff.Tag
            Call ShowData
        End If
        Me.txtStuff.SetFocus
        Call OS.PressKey(vbKeyTab)
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

Private Sub msfUnit_AfterAddRow(Row As Long)
    Dim lngCurRow As Long
    
    '修改行序号
    With msfUnit
        For lngCurRow = Row To .Rows - 1
            .TextMatrix(lngCurRow, 0) = lngCurRow
        Next
    End With
End Sub

Private Sub msfUnit_AfterDeleteRow()
    Dim lngCurRow As Long
    
    '修改行序号
    With msfUnit
        For lngCurRow = IIf(.Row <> 1, .Row - 1, .Row) To .Rows - 1
            .TextMatrix(lngCurRow, 0) = lngCurRow
        Next
    End With
End Sub

Private Sub msfUnit_CommandClick()
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    gstrSQL = "Select ID,编码,名称,简码 From 供应商 Where 末级=1 And (substr(类型,5,1)=1 Or Nvl(末级,0)=0) Order By 编码 "
    
    Call zlDatabase.OpenRecordset(rsTemp, gstrSQL, Me.Caption)
    
    With rsTemp
        If .EOF Then
            MsgBox "请初始化卫材供应商（供应商）！", vbInformation, gstrSysName
            msfUnit.SetFocus
            Exit Sub
        End If
        
        With Msf中标单位选择
            .Clear
            Set .DataSource = rsTemp
            .ColWidth(0) = 0
            .ColWidth(1) = 800
            .ColWidth(2) = 3000
            .ColWidth(3) = 800
            .Top = msfUnit.Top + msfUnit.CellTop + msfUnit.MsfObj.CellHeight
            If .Top + .Height > Me.Height Then .Top = msfUnit.Top + msfUnit.CellTop - .Height
            .Visible = True
            .ZOrder 0
            
            .Row = 1
            .ColSel = .Cols - 1
            .SetFocus
        End With
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub msf中标单位选择_DblClick()
    Dim LngFindReturn As Long, lngRow As Long, lngID As Long
    
    '先检查是否存在相同的中标单位，存在则禁止选择
    lngID = Val(Msf中标单位选择.TextMatrix(Msf中标单位选择.Row, 0))
    With msfUnit
        For lngRow = 1 To .Rows - 1
            If Val(.RowData(lngRow)) = lngID Then
                MsgBox "已经存在该中标单位，请重新选择！", vbInformation, gstrSysName
                Exit Sub
            End If
        Next
    End With
    
    With msfUnit
        .TextMatrix(.Row, 0) = .Row
        .Text = "[" & Msf中标单位选择.TextMatrix(Msf中标单位选择.Row, 1) & "]" & Msf中标单位选择.TextMatrix(Msf中标单位选择.Row, 2)
        .TextMatrix(.Row, 1) = .Text
        .RowData(.Row) = lngID
    End With
    
    With msfUnit
'        If .Row = .Rows - 1 Then
'            .Rows = .Rows + 1
'            .Row = .Row + 1
'            .TextMatrix(.Row, 0) = .Row
'        End If
        .Col = 2
        .SetFocus
    End With
End Sub

Private Sub msf中标单位选择_GotFocus()
    If Msf中标单位选择.Rows - 1 = 1 Then Call msf中标单位选择_DblClick
End Sub

Private Sub msf中标单位选择_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call msf中标单位选择_DblClick
End Sub

Private Sub msf中标单位选择_LostFocus()
    With Msf中标单位选择
        .ZOrder 1
        .Visible = False
    End With
End Sub
