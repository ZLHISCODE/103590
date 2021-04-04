VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmServiceSelect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "选择服务科室"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7410
   Icon            =   "frmServiceSelect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   7410
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picDept 
      Height          =   3795
      Left            =   2610
      ScaleHeight     =   3735
      ScaleWidth      =   4710
      TabIndex        =   1
      Top             =   15
      Width           =   4770
      Begin VSFlex8Ctl.VSFlexGrid vsfList 
         Height          =   3300
         Left            =   15
         TabIndex        =   7
         Top             =   435
         Width           =   4680
         _cx             =   8255
         _cy             =   5821
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
         BackColorSel    =   16777152
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   255
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmServiceSelect.frx":000C
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
         ExplorerBar     =   1
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
         VirtualData     =   0   'False
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
      Begin VB.TextBox txtFind 
         Height          =   300
         Left            =   600
         TabIndex        =   3
         Top             =   60
         Width           =   1335
      End
      Begin VB.CheckBox chkAllSelect 
         Caption         =   "全选"
         Height          =   255
         Left            =   2640
         TabIndex        =   2
         Top             =   83
         Width           =   720
      End
      Begin VB.Label lblFind 
         AutoSize        =   -1  'True
         Caption         =   "查找"
         Height          =   180
         Left            =   50
         TabIndex        =   4
         Top             =   120
         Width           =   360
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   4920
      TabIndex        =   6
      Top             =   3870
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6210
      TabIndex        =   5
      Top             =   3870
      Width           =   1100
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   1740
      Top             =   225
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceSelect.frx":0081
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceSelect.frx":06CD
            Key             =   "Dept"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmServiceSelect.frx":09E9
            Key             =   "Dept_No"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwDept 
      Height          =   3795
      Left            =   0
      TabIndex        =   0
      Top             =   15
      Width           =   2565
      _ExtentX        =   4524
      _ExtentY        =   6694
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   529
      LabelEdit       =   1
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "ils16"
      Appearance      =   1
   End
End
Attribute VB_Name = "frmServiceSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mintRow  As Integer  '记录当前行
Private mintFind As Integer  '用来记录查询到哪个位置了
Private mblnChkFocus As Boolean  '"全选"复选框获取焦点时为True
Private mrs部门  As Recordset    '部门记录集
Private mstrKey  As String       '选中节点的性质编码
Private mblnFind As Boolean      '通过简码(编码、名称)查询时为True
Private mblnChang As Boolean
Private mint模块 As Integer 'mint模块=1为存储库房，mint模块=2为规格批量修改存储库房

Private Sub SetColumns()
    Dim intCol As Integer
    
    With vsfList
        .Rows = 1
        .Cols = 7
        .ColDataType(0) = flexDTBoolean
        .Editable = flexEDKbdMouse
'        .ExtendLastCol = True
        .TextMatrix(0, 0) = "选择"
        .TextMatrix(0, 1) = "ID"
        .TextMatrix(0, 2) = "编码"
        .TextMatrix(0, 3) = "名称"
        .TextMatrix(0, 4) = "简码"
        .TextMatrix(0, 5) = "性质编码"
        .TextMatrix(0, 6) = "性质"
        
        For intCol = 0 To .Cols - 1
            .ColKey(intCol) = .TextMatrix(0, intCol)
        Next
        .ColWidth(.ColIndex("选择")) = 500
        .ColWidth(.ColIndex("ID")) = 0
        .ColWidth(.ColIndex("名称")) = 2000
        .ColWidth(.ColIndex("性质编码")) = 0
        .ColWidth(.ColIndex("性质")) = IIf(mblnFind, 1000, 0)
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        .ColAlignment(.ColIndex("编码")) = flexAlignLeftCenter
    End With
End Sub

Private Sub FillTree()
'获取性质赋值到树表
    Dim rs性质 As Recordset
    Dim str性质 As String
    
    gstrSql = "Select 编码, 名称 From 部门性质分类 Where Instr('3ABCDEF', 编码) > 0"
    Set rs性质 = zlDatabase.OpenSQLRecord(gstrSql, "查询性质")
    
    With tvwDept
        .Nodes.Clear
        .Nodes.Add , , "KRoot", "所有性质", "Root", "Root"
        .Nodes("KRoot").Sorted = True
        
        Do While Not rs性质.EOF
            .Nodes.Add "KRoot", tvwChild, "K" & rs性质!编码, rs性质!名称, "Dept", "Dept"
            rs性质.MoveNext
        Loop
        
        .Nodes.Item(1).Expanded = True
        .Nodes.Item(1).Selected = True
    End With
End Sub

Public Sub ShowMe(ByVal frmParent As Form, ByVal intRow As Integer, ByVal str服务对象 As String, ByVal int模块 As Integer, Optional strkey As String)
    Dim strTemp As String
    Dim strFind As String
    
    mint模块 = int模块
    mintRow = intRow
    mblnFind = (strkey <> "")
    If mblnFind Then
        strTemp = " And ( d.编码 like [2] or d.名称 like [2] or d.简码 like [2]) "
    End If
    
    gstrSql = "Select Distinct d.Id, d.编码, d.名称, d.简码, a.编码 As 性质编码, c.工作性质 as 性质 " & vbNewLine & _
            "From 部门性质分类 A, 部门性质说明 C, 部门表 D " & vbNewLine & _
            "Where d.Id = c.部门id And c.工作性质 = a.名称 And (d.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or d.撤档时间 Is Null) And " & vbNewLine & _
            "Instr('3ABCDEF', a.编码) > 0 And Instr([1], ',' || c.服务对象 || ',') > 0 " & strTemp & vbNewLine & _
            "Order By d.id,d.编码"
    Set mrs部门 = zlDatabase.OpenSQLRecord(gstrSql, "获取部门", "," & str服务对象 & ",", strkey & "%")
    
    frmServiceSelect.Show vbModal, frmParent
End Sub

Private Sub chkAllSelect_Click()
    Dim i As Integer
    
    With vsfList
        If mblnChkFocus Then
            For i = 1 To .Rows - 1
                .TextMatrix(i, .ColIndex("选择")) = IIf(chkAllSelect.Value = 1, -1, 0)
            Next
        End If
    End With
End Sub

Private Sub chkAllSelect_GotFocus()
    mblnChkFocus = True
End Sub

Private Sub chkAllSelect_LostFocus()
    mblnChkFocus = False
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    If mblnFind Then
        Me.Width = 6000
        tvwDept.Visible = False
    Else
        Call FillTree
    End If
    
    Call SetColumns
    Call FillList("KRoot")
    mintFind = 0
End Sub

Private Sub cmdOK_Click()
    Dim blnCancel As Boolean
    Dim lngRow As Long, lngRows As Long
    Dim str科室 As String, str科室ID As String
    
    '循环提取用户所选择的科室
    With vsfList
        For lngRow = 1 To .Rows - 1
            If Val(.TextMatrix(lngRow, .ColIndex("选择"))) = -1 Then
                str科室 = str科室 & "," & .TextMatrix(lngRow, .ColIndex("名称"))
                str科室ID = str科室ID & "," & .TextMatrix(lngRow, .ColIndex("ID"))
            End If
        Next
    End With
    
    If str科室 <> "" Then
        str科室 = Mid(str科室, 2)
        str科室ID = Mid(str科室ID, 2)
    End If
    If mint模块 = 1 Then '存储库房模块
        With frmServiceSectOffice.msfServiceSectOffice
            If str科室 <> "" Then .TextMatrix(.Row, 1) = "√"
            .Text = str科室
            .TextMatrix(mintRow, 3) = .Text
            .TextMatrix(mintRow, 4) = str科室ID
            If .Rows - 1 > .Row Then .Row = .Row + 1
        End With
    Else   '规格批量修改模块
        With frmServiceDepartment.vsfDepartment
            .TextMatrix(.Row, .Col) = str科室
            .TextMatrix(.Row, .Col + 1) = str科室ID
        End With
    End If
    
    Unload Me
End Sub

Private Sub Form_Resize()
    If mblnFind Then
        picDept.Move 0, picDept.Top, Me.ScaleWidth
        cmdCancel.Move cmdCancel.Left - 800
        cmdOK.Move cmdOK.Left - 800
        vsfList.Width = picDept.ScaleWidth - 10
    End If
End Sub

Private Sub vsfList_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfList
        If Col <> 0 Then
            Cancel = True
        End If
        mblnChkFocus = False
    End With
End Sub

Private Sub vsfList_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Dim intRow As Integer
    Dim intCount As Integer
    
    With vsfList
        If Row > 0 And Col = 0 Then
            For intRow = 1 To .Rows - 1
                If Val(.TextMatrix(intRow, .ColIndex("选择"))) = -1 Then
                    intCount = intCount + 1
                End If
            Next
            
            '是否全选
            If mblnChkFocus = False Then
                If intCount = .Rows - 1 Then
                    chkAllSelect.Value = 1
                Else
                    chkAllSelect.Value = 0
                End If
            End If
        End If
    End With
End Sub

Private Sub vsfList_ChangeEdit()
    mblnChang = True
End Sub

Private Sub tvwDept_NodeClick(ByVal Node As MSComctlLib.Node)
    If mstrKey = Mid(Node.Key, 2) Then Exit Sub
    Call FillList(Mid(Node.Key, 2))
End Sub

Private Sub FillList(Optional ByVal strkey As String)
'获取科室信息
    Dim i As Integer
    Dim intCount As Integer
    
    With vsfList
        mstrKey = strkey
        If strkey = "KRoot" Then
            .Rows = 1
            Do While Not mrs部门.EOF
                For i = 1 To .Rows - 1
                    If mrs部门!ID = .TextMatrix(i, .ColIndex("ID")) Then
                        mrs部门.MoveNext
                        If mrs部门.EOF Then
                            chkAllSelect.Value = IIf(intCount = .Rows - 1, 1, 0)
                            .Row = 1
                            Exit Sub
                        End If
                        i = 0
                    End If
                Next
                
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, .ColIndex("ID")) = IIf(IsNull(mrs部门!ID), "", mrs部门!ID)
                .TextMatrix(.Rows - 1, .ColIndex("编码")) = IIf(IsNull(mrs部门!编码), "", mrs部门!编码)
                .TextMatrix(.Rows - 1, .ColIndex("名称")) = IIf(IsNull(mrs部门!名称), "", mrs部门!名称)
                .TextMatrix(.Rows - 1, .ColIndex("简码")) = IIf(IsNull(mrs部门!简码), "", mrs部门!简码)
                .TextMatrix(.Rows - 1, .ColIndex("性质编码")) = IIf(IsNull(mrs部门!性质编码), "", mrs部门!性质编码)
                .TextMatrix(.Rows - 1, .ColIndex("性质")) = IIf(IsNull(mrs部门!性质), "", mrs部门!性质)
    
                If mint模块 = 1 Then '存储库房模块
                    If InStr(1, "," & frmServiceSectOffice.msfServiceSectOffice.TextMatrix(frmServiceSectOffice.msfServiceSectOffice.Row, 4) & ",", "," & mrs部门!ID & ",") > 0 Then
                        .TextMatrix(.Rows - 1, .ColIndex("选择")) = -1
                        intCount = intCount + 1
                    End If
                Else    '规格批量修改库房模块
                    If InStr(1, "," & frmServiceDepartment.vsfDepartment.TextMatrix(frmServiceDepartment.vsfDepartment.Row, 3) & ",", "," & mrs部门!ID & ",") > 0 Then
                        .TextMatrix(.Rows - 1, .ColIndex("选择")) = -1
                        intCount = intCount + 1
                    End If
                End If
                
                mrs部门.MoveNext
            Loop
            chkAllSelect.Value = IIf(intCount = .Rows - 1, 1, 0)
            .Row = 1
        Else
            For i = 1 To .Rows - 1
                .RowHidden(i) = False
                If .TextMatrix(i, .ColIndex("性质编码")) <> strkey And strkey <> "Root" Then
                    .RowHidden(i) = True
                End If
            Next
        End If
    End With
End Sub

Private Sub txtFind_GotFocus()
    zlControl.TxtSelAll txtFind
End Sub

Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strFind As String
    Dim i As Integer
    Dim blnResult As Boolean
    Dim j As Integer
    Dim k As Integer
    
    blnResult = False
    With vsfList
        If KeyCode = vbKeyReturn And Trim(txtFind.Text) <> "" Then
            strFind = UCase(Trim(txtFind.Text))
            If mintFind > .Rows - 1 Then
                mintFind = 1
            Else
                mintFind = mintFind + 1
                If mintFind > .Rows - 1 Then
                    mintFind = 1
                End If
            End If
            
            For i = mintFind To .Rows - 1
                If IsNumeric(strFind) Then
                    If .TextMatrix(i, .ColIndex("编码")) = strFind Then
                        .Row = i
                        .TopRow = i
                        mintFind = i
                        Call SelectNode
                        Exit Sub
                    End If
                    
                    If i = .Rows - 1 Then
                        For k = 1 To mintFind
                            If .TextMatrix(k, .ColIndex("编码")) = strFind Then
                                .Row = k
                                .TopRow = k
                                mintFind = k
                                Call SelectNode
                                Exit Sub
                            End If
                        Next
                    End If
                Else
                    If .TextMatrix(i, .ColIndex("简码")) Like "*" & strFind & "*" Then
                        .Row = i
                        .TopRow = i
                        mintFind = i
                        Call SelectNode
                        blnResult = True
                        Exit Sub
                    End If
                    
                    If i = .Rows - 1 Then
                        For k = 1 To mintFind
                            If .TextMatrix(k, .ColIndex("简码")) Like "*" & strFind & "*" Then
                                .Row = k
                                .TopRow = k
                                mintFind = k
                                Call SelectNode
                                blnResult = True
                                Exit Sub
                            End If
                        Next
                    End If
                End If
            Next
            
            If blnResult = False Then
                For i = mintFind To .Rows - 1
                    If .TextMatrix(i, .ColIndex("名称")) Like "*" & strFind & "*" Then
                        .Row = i
                        .TopRow = i
                        mintFind = i
                        Call SelectNode
                        blnResult = True
                        Exit Sub
                    End If
                Next
                
                For k = 1 To mintFind
                    If .TextMatrix(k, .ColIndex("名称")) Like "*" & strFind & "*" Then
                        .Row = k
                        .TopRow = k
                        mintFind = k
                        Call SelectNode
                        blnResult = True
                        Exit Sub
                    End If
                Next
            End If
        End If
    End With
End Sub

Private Sub SelectNode()
'确定查找部门的性质
    Dim i As Integer
    
    With vsfList
        For i = 1 To tvwDept.Nodes.Count
            If .TextMatrix(.Row, .ColIndex("性质编码")) = Mid(tvwDept.Nodes(i).Key, 2) Then
                tvwDept.Nodes(i).Selected = True
                Call FillList(Mid(tvwDept.Nodes(i).Key, 2))
                Exit Sub
            End If
        Next
    End With
End Sub


