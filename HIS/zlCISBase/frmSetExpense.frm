VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSetExpense 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "费别设置"
   ClientHeight    =   8235
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11280
   Icon            =   "frmSetExpense.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8235
   ScaleWidth      =   11280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox pic药品 
      Height          =   5415
      Left            =   120
      ScaleHeight     =   5355
      ScaleWidth      =   3435
      TabIndex        =   18
      Top             =   2040
      Width           =   3495
      Begin MSComctlLib.TreeView tvwDetails 
         Height          =   3480
         Left            =   0
         TabIndex        =   19
         Tag             =   "1000"
         Top             =   240
         Width           =   3240
         _ExtentX        =   5715
         _ExtentY        =   6138
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   494
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         FullRowSelect   =   -1  'True
         ImageList       =   "ImgTvw"
         Appearance      =   0
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "关闭(&C)"
      Height          =   350
      Left            =   10080
      TabIndex        =   17
      Top             =   7680
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "保存(&O)"
      Height          =   350
      Left            =   8760
      TabIndex        =   16
      Top             =   7680
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   120
      TabIndex        =   15
      Top             =   7680
      Width           =   1100
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "清除(&D)"
      Height          =   350
      Left            =   3840
      TabIndex        =   14
      Top             =   7680
      Visible         =   0   'False
      Width           =   1100
   End
   Begin VB.Frame fra药品应用 
      Caption         =   "应用范围"
      Height          =   1620
      Left            =   3840
      TabIndex        =   7
      Top             =   240
      Width           =   7335
      Begin VB.OptionButton opt应用于 
         Caption         =   "应用于本分类的所有药品(&5)"
         Height          =   255
         Index           =   5
         Left            =   3240
         TabIndex        =   13
         Top             =   1080
         Width           =   2955
      End
      Begin VB.OptionButton opt应用于 
         Caption         =   "应用于同级的所有药品(&4)"
         Height          =   255
         Index           =   4
         Left            =   3240
         TabIndex        =   12
         Top             =   720
         Width           =   3075
      End
      Begin VB.OptionButton opt应用于 
         Caption         =   "应用于本品种下所有药品(&1)"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   11
         Top             =   720
         Width           =   2835
      End
      Begin VB.OptionButton opt应用于 
         Caption         =   "应用于所有“片剂”类药品(&3)"
         Height          =   255
         Index           =   3
         Left            =   3240
         TabIndex        =   10
         Top             =   360
         Width           =   3915
      End
      Begin VB.OptionButton opt应用于 
         Caption         =   "应用于所有“西成药”(&2)"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   9
         Top             =   1080
         Width           =   2715
      End
      Begin VB.OptionButton opt应用于 
         Caption         =   "仅应用于本规格药品(&0)"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Value           =   -1  'True
         Width           =   2595
      End
   End
   Begin VB.ComboBox cbo计算方法 
      Height          =   300
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   960
      Width           =   2535
   End
   Begin VB.ComboBox cbo费别 
      Height          =   300
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   360
      Width           =   2535
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfList 
      Height          =   4920
      Left            =   3840
      TabIndex        =   3
      Top             =   2400
      Width           =   7320
      _cx             =   12912
      _cy             =   8678
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
      BackColorSel    =   16769992
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   10329501
      GridColorFixed  =   10329501
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSetExpense.frx":000C
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
      Editable        =   0
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
   Begin MSComctlLib.ImageList ImgTvw 
      Left            =   2640
      Top             =   1200
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
            Picture         =   "frmSetExpense.frx":0083
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetExpense.frx":061D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSetExpense.frx":6E7F
            Key             =   "规格U"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label5 
      Caption         =   "药品品种"
      Height          =   180
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label lblNote 
      Caption         =   "每一收入项目可按应收金额划分为多段(最多16段)，设置不同的实收比例。"
      Height          =   180
      Left            =   3840
      TabIndex        =   2
      Top             =   2040
      Width           =   6735
   End
   Begin VB.Label lblMeasure 
      Caption         =   "计算方法"
      Height          =   180
      Left            =   120
      TabIndex        =   1
      Top             =   1020
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "选择费别"
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   420
      Width           =   855
   End
End
Attribute VB_Name = "frmSetExpense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngId As Long            '药品id
Private mstrGrade As String

Public Function ShowMe(objfrm As Object, ByVal lngId As Long, ByVal strgrade As String) As Boolean
    mlngId = lngId
    mstrGrade = strgrade
    Me.Show vbModal, objfrm
    
End Function
'
Private Sub LoadCharge()
    Dim rsTemp As ADODB.Recordset
    Dim intIndex As Integer

    gstrSql = "Select 名称 From 费别 Order By 编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "取费别")

    cbo费别.Clear

    With rsTemp
        Do While Not .EOF
            cbo费别.AddItem !名称

            If !名称 = mstrGrade Then
                intIndex = cbo费别.ListCount - 1
            End If

            .MoveNext
        Loop
    End With

    If cbo费别.ListCount > 0 Then
        cbo费别.ListIndex = intIndex
    End If

End Sub

Private Sub cbo费别_Click()
    If tvwDetails.SelectedItem Is Nothing Then Exit Sub
    If tvwDetails.SelectedItem.Tag Like "*品种" Or tvwDetails.SelectedItem.Tag Like "*分类" Then
        vsfList.Rows = 1
        Exit Sub
    End If
    
    If tvwDetails.SelectedItem Is Nothing Then Exit Sub
    If tvwDetails.SelectedItem.Tag Like "*品种" Or tvwDetails.SelectedItem.Tag Like "*分类" Then
        vsfList.Rows = 1
        Exit Sub
    End If
    Call LoadChargeList(Mid(tvwDetails.SelectedItem.Key, 3))
End Sub

Private Sub cbo计算方法_Click()
'    1-成本价加收比例计算,不分段
    If cbo计算方法.ListIndex = 1 Then
        lblNote.Caption = "  药品实收金额=成本价*(1+加收比率)，如果不是药品将忽略此设置，不打折。"
        vsfList.TextMatrix(0, 1) = "分段起点"
        vsfList.TextMatrix(0, 2) = "加收比率(%)"
    '0-分段比例计算
    Else
       lblNote.Caption = "    每一收入项目可按应收金额划分为多段(最多16段)，设置不同的实收比例。"
       vsfList.TextMatrix(0, 1) = "应收分段起点"
       vsfList.TextMatrix(0, 2) = "实收比率(%)"
    End If
    
    If tvwDetails.SelectedItem Is Nothing Then Exit Sub
    If tvwDetails.SelectedItem.Tag Like "*品种" Or tvwDetails.SelectedItem.Tag Like "*分类" Then
        vsfList.Rows = 1
        Exit Sub
    End If
    Call LoadChargeList(Mid(tvwDetails.SelectedItem.Key, 3))
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdOK_Click()
    Dim i As Integer
    
    If tvwDetails.SelectedItem.Tag Like "*分类" Or tvwDetails.SelectedItem.Tag Like "*品种" Then Exit Sub
    If CheckData = True Then
        If opt应用于(0).Value = False Then
            For i = 1 To opt应用于.UBound
                If opt应用于(i).Value = True Then
                    If MsgBox("该药品的费别设置应用范围为“" & opt应用于(i).Caption & "”是否继续？", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Sub
                    Else
                        Exit For
                    End If
                End If
            Next
        End If
        Call SaveCharge
    End If
End Sub

Private Sub Form_Load()
    Dim objNode As Node
    
    'ByZT20030722
    If glngSys Like "8??" Then
        Caption = "会员等级单项收费设置"
    End If
    
    '计算方法
    cbo计算方法.AddItem "0-分段比例计算", 0
    cbo计算方法.AddItem "1-成本价加收比例计算", 1
    cbo计算方法.ListIndex = 0
    
    '取费别
    Call LoadCharge
    '填充树
    Call FullTreeView
    Call InitVsf    '初始化控件
    
    Call LoadChargeList(mlngId)
    For Each objNode In tvwDetails.Nodes
        If Mid(objNode.Key, 3) = mlngId Then
            objNode.Selected = True
            objNode.Expanded = True
        End If
    Next
End Sub


Private Function CheckData() As Boolean
    '检查数据不能有空行
    Dim intRow As Integer
    Dim intCol As Integer
    With vsfList
        For intRow = 1 To .Rows - 1
            For intCol = 1 To .Cols - 1
                If Trim(.TextMatrix(intRow, intCol)) = "" Then
                    MsgBox "单元格不能为空！", vbInformation, gstrSysName
                    CheckData = False
                    vsfList.SetFocus
                    .Row = intRow
                    .Col = intCol
                    Exit Function
                End If
            Next
        Next
        CheckData = True
    End With
End Function
Private Sub LoadChargeList(ByVal lngId As Long)
    Dim rsTemp As ADODB.Recordset
    Dim i As Long
    Dim strSQL As String

    gstrSql = "Select 段号, 应收段首值, 应收段尾值, 实收比率, 计算方法 " & _
        " From 费别明细 Where 费别 = [1] And 收费细目id=[2] And 计算方法=[3]" & strSQL & " Order By 段号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "取费别明细", cbo费别.Text, lngId, cbo计算方法.ListIndex)

    vsfList.Rows = 2
    If rsTemp.RecordCount = 0 Then
        vsfList.TextMatrix(1, 0) = 1
        vsfList.TextMatrix(1, 1) = "0.00"
        vsfList.TextMatrix(1, 2) = "100.00"
        Exit Sub
    End If

    cbo计算方法.ListIndex = IIf(rsTemp!计算方法 = 0, 0, 1)

    With rsTemp
        vsfList.Rows = .RecordCount + 1
        cbo计算方法.ListIndex = Val(.Fields("计算方法").Value)     '调用Click事件设置相关控件

        For i = 1 To .RecordCount
            If i > 16 Then Exit For
            vsfList.TextMatrix(i, 0) = i
            vsfList.TextMatrix(i, 1) = Format(.Fields("应收段首值").Value, "###########0.00;-##########0.00;0.00;0.00")
            vsfList.TextMatrix(i, 2) = Format(.Fields("实收比率").Value, "###0.000;-##0.000;0.000;0.000")
            .MoveNext
        Next
    End With
    With vsfList
        .Cell(flexcpBackColor, 1, 1, 1, 1) = &H8000000F
    End With
End Sub

Private Sub FullTreeView()
    Dim NodeThis As Node
    Dim Int末级 As Integer
    Dim lng库房ID As Long
    Dim rs材质分类 As ADODB.Recordset
    Dim recdate As ADODB.Recordset
    
    '药品用途分类是否有数据
    gstrSql = " Select 编码,名称 From 诊疗项目类别 " & _
              " Where Instr([1],编码,1) > 0 " & _
              " Order by 编码"
    Set rs材质分类 = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, "567")
    
    If rs材质分类.RecordCount = 0 Then
        Exit Sub
    End If
    
    With tvwDetails
        .Nodes.Clear
        Do While Not rs材质分类.EOF
            .Nodes.Add , , "Root" & rs材质分类!名称, rs材质分类!名称, 1, 1
            .Nodes("Root" & rs材质分类!名称).Tag = rs材质分类!编码
            rs材质分类.MoveNext
        Loop
    End With
    '分类
    gstrSql = "Select ID, 上级id, 编码, 名称, Decode(类型, 1, '西成药', 2, '中成药', 3, '中草药') 分类, '分类' As 类别" & vbNewLine & _
                "From 诊疗分类目录" & vbNewLine & _
                "Where 类型 In ('1', '2', '3') And Nvl(To_Char(撤档时间, 'YYYY-MM-DD'), '3000-01-01') = '3000-01-01'" & vbNewLine & _
                "Start With 上级id Is Null" & vbNewLine & _
                "Connect By Prior ID = 上级id"
    Set recdate = zlDatabase.OpenSQLRecord(gstrSql, "分类")
    
    If recdate.EOF Then
        MsgBox "请初始化药品用途分类（药品用途分类）！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    With recdate
        Do While Not .EOF
            If IsNull(!上级ID) Then
                Set NodeThis = tvwDetails.Nodes.Add("Root" & !分类, 4, "分类K_" & !ID, !名称, 1, 1)
            Else
                Set NodeThis = tvwDetails.Nodes.Add("分类K_" & !上级ID, 4, "分类K_" & !ID, !名称, 1, 1)
            End If
            NodeThis.Tag = !分类 & "-" & !类别  '存放分类类型:1-西成药,2-中成药,3-中草药
            .MoveNext
        Loop
    End With
    '品种
    gstrSql = "Select ID, 分类id, 编码, 名称, Decode(类别, 5, '西成药', 6, '中成药', 7, '中草药') 分类, '品种' As 类别" & vbNewLine & _
                "From 诊疗项目目录" & vbNewLine & _
                "Where 分类id In (Select ID" & vbNewLine & _
                "               From 诊疗分类目录" & vbNewLine & _
                "               Where 类型 In ('1', '2', '3') And Nvl(To_Char(撤档时间, 'YYYY-MM-DD'), '3000-01-01') = '3000-01-01'" & vbNewLine & _
                "               Start With 上级id Is Null" & vbNewLine & _
                "               Connect By Prior ID = 上级id)"
    Set recdate = zlDatabase.OpenSQLRecord(gstrSql, "品种")
    With recdate
        Do While Not .EOF
            Set NodeThis = tvwDetails.Nodes.Add("分类K_" & !分类id, 4, "K_" & !ID, !名称, 1, 1)
            NodeThis.Tag = !分类 & "-" & !类别  '存放分类类型:1-西成药,2-中成药,3-中草药
            .MoveNext
        Loop
    End With
    '规格
    gstrSql = "Select a.药品id As ID, a.药名id As 上级id, b.编码,b.规格 as 名称, b.分类, b.类别" & _
               " From 药品规格 A," & _
                "     (Select ID, 分类id, 编码, 规格, Decode(类别, '5', '西成药', '6', '中成药', '7', '中草药') 分类, '药品' As 类别" & _
                      " From 收费项目目录" & _
                      " Where 类别 In ('5', '6', '7') And Nvl(To_Char(撤档时间, 'YYYY-MM-DD'), '3000-01-01') = '3000-01-01') B" & _
               " Where a.药品id = b.Id"
    Set recdate = zlDatabase.OpenSQLRecord(gstrSql, "规格查询")
    
    With recdate
        Do While Not .EOF
            Set NodeThis = tvwDetails.Nodes.Add("K_" & !上级ID, 4, "M_" & !ID, IIf(IsNull(!名称), "", !名称), 1, 1)
            NodeThis.Tag = !分类 & "-" & !类别  '存放分类类型:1-西成药,2-中成药,3-中草药
            .MoveNext
        Loop
    End With
    
    With tvwDetails
        If .Nodes.Count <> 0 Then
            .Nodes(1).Selected = True
            If .Nodes(1).Children <> 0 Then
                Int末级 = 1
                .Nodes(Int末级).Child.Selected = True
                .SelectedItem.Selected = True
            ElseIf .Nodes(2).Children <> 0 Then
                Int末级 = 2
                .Nodes(Int末级).Child.Selected = True
                .SelectedItem.Selected = True
            ElseIf .Nodes(3).Children <> 0 Then
                Int末级 = 3
                .Nodes(Int末级).Child.Selected = True
                .SelectedItem.Selected = True
            Else
                Int末级 = 0
                .Nodes(1).Selected = True
                .SelectedItem.Selected = True
            End If
            If Int末级 <> 0 Then .Nodes(Int末级).Expanded = True
        End If
    End With
    tvwDetails.Move 0, 0, pic药品.Width, pic药品.Height
End Sub

Private Sub InitVsf()
    '初始化vsflexgrid
    With vsfList
        .Cols = 3
        .Rows = 1
        .Editable = flexEDNone
        .SelectionMode = flexSelectionFree
        
        .TextMatrix(0, 0) = "分段号"
        If cbo计算方法.Text = "0-分段比例计算" Then
            .TextMatrix(0, 1) = "应收分段起点"
            .TextMatrix(0, 2) = "加收比率(%)"
        ElseIf cbo计算方法.Text = "1-成本价加收比例计算" Then
            .TextMatrix(0, 1) = "分段起点"
            .TextMatrix(0, 2) = "加收比率(%)"
        End If
    End With
End Sub

Private Sub opt应用于_Click(Index As Integer)
    Dim i As Integer
    
    For i = 1 To opt应用于.UBound
        If i = Index Then
            opt应用于(i).FontBold = True
        Else
            opt应用于(i).FontBold = False
        End If
    Next
End Sub

Private Sub tvwDetails_NodeClick(ByVal Node As MSComctlLib.Node)
    With tvwDetails
        If Node.Tag Like "*分类" = True Or Node.Tag Like "*品种" Or Node.Tag = "5" Or Node.Tag = "6" Or Node.Tag = "7" Then
            vsfList.Rows = 1
            vsfList.Editable = flexEDNone
            Exit Sub
        End If
        Call GetDrugOtherInfo(Mid(Node.Key, 3))
        Call LoadChargeList(Mid(tvwDetails.SelectedItem.Key, 3))
        vsfList.SetFocus
    End With
End Sub

Private Sub GetDrugOtherInfo(ByVal lngItemId As Long)
    '主要用于药品目录管理中得到当前药品的剂型和材质
    Dim rsTemp As ADODB.Recordset
    Dim str材质 As String
    If lngItemId = 0 Then Exit Sub
    
    gstrSql = "Select Decode(A.类别, '5', '西成药', '6', '中成药', '中草药') As 类别, B.药品剂型 " & _
        " From 收费项目目录 A, 药品特性 B, 药品规格 C " & _
        " Where A.ID = C.药品id And B.药名id = C.药名id And A.ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "取药品信息", lngItemId)
    
    If Not rsTemp.EOF Then
        opt应用于(2).Caption = "应用于所有“" & rsTemp!类别 & "”(&2)"
        opt应用于(3).Caption = "应用于所有“" & rsTemp!药品剂型 & "”类药品(&3)"
    End If
End Sub

Private Sub vsfList_DblClick()
    With vsfList
        If .Editable = flexEDKbdMouse Then
            .EditCell
            .EditSelStart = 0
            .EditSelLength = Len(.EditText)
        End If
    End With
End Sub

Private Sub vsfList_EnterCell()
    With vsfList
        If .Col = 1 And .Row = 1 Then
            .Editable = flexEDNone
        Else
            .Editable = flexEDKbdMouse
        End If
    End With
End Sub

Private Sub vsfList_KeyDown(KeyCode As Integer, Shift As Integer)
    With vsfList
        If KeyCode = vbKeyReturn Then
            If .TextMatrix(.Row, .Col) = "" Then
                MsgBox "不能为空！", vbInformation, gstrSysName
                vsfList.SetFocus
                KeyCode = 0
                Exit Sub
            End If
            If .Col <> .Cols - 1 Then
                .Col = .Col + 1
            ElseIf .Col < 16 And cbo计算方法.ListIndex <> 1 Then
                .Rows = .Rows + 1
                .Row = .Row + 1
                .TextMatrix(.Row, 0) = .Row
                .Col = 1
            End If
        ElseIf KeyCode = vbKeyDelete Then
            If .Rows > 2 Then
                .RemoveItem .Row
            Else
                .TextMatrix(1, 0) = ""
                .TextMatrix(1, 1) = ""
                .TextMatrix(1, 2) = ""
            End If
        End If
    End With
End Sub

Private Sub vsfList_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    With vsfList
        If IsNumeric(Chr(KeyAscii)) = False And (KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete) And Chr(KeyAscii) <> "." And KeyAscii <> vbKeyReturn Then
            KeyAscii = 0
        End If
    End With
End Sub

Private Sub vsfList_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfList
        If IsNumeric(.EditText) = False Then
            MsgBox "请输入正确的数字", vbInformation, gstrSysName
            Cancel = True
            Exit Sub
        End If
        If Row <> 1 And Val(.EditText) < Val(.TextMatrix(Row - 1, 1)) And Col = 1 Then
            MsgBox "应手段值必须由小到大！", vbInformation, gstrSysName
            Cancel = True
            Exit Sub
        End If
        If Row <> 1 And Val(.EditText) = Val(.TextMatrix(Row - 1, 1)) And Col = 2 Then
            MsgBox "相邻行比率不能相同！", vbInformation, gstrSysName
            Cancel = True
            Exit Sub
        End If
    End With
End Sub


Private Sub SaveCharge()
    Dim str比率 As String
    Dim curStart As Currency, curEnd As Currency, dblTax As Double
    Dim intRow As Long
    Dim blnTrans As Boolean
    Dim int应用 As Integer
    Dim lngId As Long
    
    On Error GoTo ErrHand
    lngId = Mid(tvwDetails.SelectedItem.Key, 3)
    If vsfList.Rows = 1 Then Exit Sub
'    If vsfList.Editable = flexEDNone Then Exit Sub
    With vsfList
        For intRow = 1 To .Rows - 1
            curStart = Val(.TextMatrix(intRow, 0))
            If intRow >= .Rows - 1 Then
                curEnd = Val("10000000000.00")
            Else
                curEnd = Val(.TextMatrix(intRow, 1)) - 0.01
            End If
            dblTax = .TextMatrix(intRow, 2)
            str比率 = str比率 & intRow & ":" & curStart & ":" & curEnd & ":" & dblTax & ";"
        Next
    End With
    
    gcnOracle.BeginTrans

    '药品目录中设置费别
    If opt应用于(0).Value = True Then
        int应用 = 0
    ElseIf opt应用于(1).Value = True Then
        int应用 = 1
    ElseIf opt应用于(2).Value = True Then
        int应用 = 2
    ElseIf opt应用于(3).Value = True Then
        int应用 = 3
    ElseIf opt应用于(4).Value = True Then
        int应用 = 4
    ElseIf opt应用于(5).Value = True Then
        int应用 = 5
    End If

    gstrSql = "zl_费别明细_update('" & cbo费别.Text & "'," & lngId & ",'" & str比率 & "'," & Val(cbo计算方法.Text) & "," & 3 & "," & int应用 & ")"
    Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)

    gcnOracle.CommitTrans
    MsgBox "保存成功！", vbInformation, gstrSysName
    Exit Sub
ErrHand:
    If Not blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
End Sub

