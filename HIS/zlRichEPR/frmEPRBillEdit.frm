VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmEPRBillEdit 
   BorderStyle     =   0  'None
   Caption         =   "诊疗单据编辑"
   ClientHeight    =   4350
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   6285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VSFlex8Ctl.VSFlexGrid vfg附项 
      Height          =   1230
      Left            =   150
      TabIndex        =   8
      Top             =   1440
      Width           =   6000
      _cx             =   10583
      _cy             =   2170
      Appearance      =   2
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
      BackColorSel    =   16635590
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   4
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
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
   Begin VB.PictureBox picEdit 
      Align           =   1  'Align Top
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   3795
      Left            =   0
      ScaleHeight     =   3795
      ScaleWidth      =   6285
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   6285
      Begin VB.PictureBox picApply 
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   1850
         Left            =   120
         ScaleHeight     =   1845
         ScaleWidth      =   6135
         TabIndex        =   27
         Top             =   1200
         Visible         =   0   'False
         Width           =   6135
         Begin VSFlex8Ctl.VSFlexGrid vsFile 
            Height          =   1470
            Left            =   35
            TabIndex        =   29
            Top             =   245
            Width           =   6000
            _cx             =   10583
            _cy             =   2593
            Appearance      =   2
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
            BackColorSel    =   16635590
            ForeColorSel    =   -2147483640
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483636
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   0
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   4
            Cols            =   3
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   360
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   ""
            ScrollTrack     =   -1  'True
            ScrollBars      =   2
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
         Begin VB.Label lblApplysMark 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "自定义申请单: (需要设置对应的xsl格式文件、xml数据文件、html显示文件)"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   30
            TabIndex        =   28
            Top             =   0
            Width           =   6120
         End
      End
      Begin VB.PictureBox picApplyType 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         ScaleHeight     =   285
         ScaleWidth      =   4035
         TabIndex        =   24
         Top             =   888
         Width           =   4035
         Begin VB.OptionButton optApply 
            Caption         =   "申请附项模式"
            Height          =   210
            Index           =   0
            Left            =   45
            TabIndex        =   26
            Top             =   45
            Value           =   -1  'True
            Width           =   1530
         End
         Begin VB.OptionButton optApply 
            Caption         =   "自定义申请单"
            Height          =   210
            Index           =   1
            Left            =   1680
            TabIndex        =   25
            Top             =   45
            Width           =   1770
         End
      End
      Begin VB.ComboBox cbxSubClass 
         Height          =   300
         Left            =   4755
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   135
         Width           =   1395
      End
      Begin VB.Frame frasplit 
         BackColor       =   &H00FFC0C0&
         Height          =   30
         Left            =   105
         TabIndex        =   21
         Top             =   3405
         Width           =   6120
      End
      Begin VB.PictureBox picEditType 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2100
         ScaleHeight     =   285
         ScaleWidth      =   4035
         TabIndex        =   17
         Top             =   3090
         Width           =   4035
         Begin VB.OptionButton optEditType 
            Caption         =   "表格病历编辑器"
            Height          =   210
            Index           =   1
            Left            =   2160
            TabIndex        =   19
            Top             =   45
            Width           =   1770
         End
         Begin VB.OptionButton optEditType 
            Caption         =   "全文病历编辑器"
            Height          =   210
            Index           =   0
            Left            =   45
            TabIndex        =   18
            Top             =   45
            Width           =   1770
         End
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "下移一行(&N)"
         Height          =   350
         Index           =   3
         Left            =   4905
         TabIndex        =   12
         Top             =   2685
         Width           =   1245
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "删除行(&D)"
         Height          =   350
         Index           =   1
         Left            =   1395
         TabIndex        =   10
         Top             =   2685
         Width           =   1245
      End
      Begin VB.TextBox txt名称 
         Height          =   300
         Left            =   1995
         MaxLength       =   60
         TabIndex        =   4
         Top             =   135
         Width           =   2115
      End
      Begin VB.TextBox txt编号 
         Height          =   300
         Left            =   600
         MaxLength       =   13
         TabIndex        =   2
         Top             =   135
         Width           =   780
      End
      Begin VB.TextBox txt说明 
         Height          =   300
         Left            =   600
         MaxLength       =   60
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   540
         Width           =   5550
      End
      Begin VB.CheckBox chk报告 
         Caption         =   "执行后有执行报告:"
         Height          =   195
         Left            =   150
         TabIndex        =   13
         Top             =   3135
         Width           =   1845
      End
      Begin VB.OptionButton opt报告 
         Caption         =   "按编辑格式打印(&1)"
         Height          =   180
         Index           =   0
         Left            =   2145
         TabIndex        =   14
         Top             =   3525
         Width           =   1890
      End
      Begin VB.OptionButton opt报告 
         Caption         =   "自定义报表打印(&2)"
         Height          =   180
         Index           =   1
         Left            =   4260
         TabIndex        =   15
         Top             =   3525
         Width           =   1890
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "添加行(&A)"
         Height          =   350
         Index           =   0
         Left            =   150
         TabIndex        =   9
         Top             =   2685
         Width           =   1245
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "上移一行(&U)"
         Height          =   350
         Index           =   2
         Left            =   3660
         TabIndex        =   11
         Top             =   2685
         Width           =   1245
      End
      Begin MSComDlg.CommonDialog cdgFile 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "分类"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4320
         TabIndex        =   22
         Top             =   195
         Width           =   360
      End
      Begin VB.Label Label1 
         Caption         =   "全文病历编辑打印方式"
         Height          =   255
         Left            =   150
         TabIndex        =   20
         Top             =   3495
         Width           =   1845
      End
      Begin VB.Label lbl名称 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "名称"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   1545
         TabIndex        =   3
         Top             =   195
         Width           =   360
      End
      Begin VB.Label lbl编号 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "编码"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   150
         TabIndex        =   1
         Top             =   195
         Width           =   360
      End
      Begin VB.Label lbl说明 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "说明"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   150
         TabIndex        =   5
         Top             =   600
         Width           =   360
      End
      Begin VB.Label lbl附项 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "申请附项: (需临床医嘱申请时额外填写的附加内容)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   150
         TabIndex        =   7
         Top             =   1215
         Width           =   4140
      End
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "单据申请和采用自定义报表输出的报告，需要管理员在自定义报表工具中，对报表'ZLCISBILL00编号-?'进行修改设计调整。"
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   570
      TabIndex        =   16
      Top             =   3870
      Width           =   5610
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgNote 
      Height          =   240
      Left            =   150
      Picture         =   "frmEPRBillEdit.frx":0000
      Top             =   3840
      Width           =   240
   End
End
Attribute VB_Name = "frmEPRBillEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum mCol
    项目 = 0: 必填: 只读: 要素id: 内容
End Enum
Private mlngBillID As Long          '当前显示的项目id
Private mstrCombos As String        '要素列表
Private mrsItems As New Recordset
Private mobjFile As New FileSystemObject     '文件操作对象
Private arrSQL() As String
Private arrSQLFile() As String
Private mblndb As Boolean
'--------------------------------------------
'以下为窗体公共方法
'--------------------------------------------

Public Function zlRefresh(lngBillId As Long) As Boolean
    '功能：根据项目id刷新当前显示内容
Dim rsTemp As New ADODB.Recordset
Dim i As Long
Dim lngCount As Long
    
    mlngBillID = lngBillId
    ReDim arrSQL(0): ReDim arrSQLFile(0)
    
    '清除此前项目的显示
    Me.txt编号.Text = "": Me.txt名称.Text = "": Me.txt说明.Text = ""
    Me.chk报告.Value = vbUnchecked
    
    '获取指定项目的信息
    Err = 0: On Error GoTo errHand
    gstrSQL = "Select ID, 编号, 名称, 说明,保留, 通用,子类, 格式 From 病历文件列表 Where 种类 = 7 And ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngBillId)
    With rsTemp
        Me.txt编号.MaxLength = .Fields("编号").DefinedSize
        Me.txt名称.MaxLength = .Fields("名称").DefinedSize
        Me.txt说明.MaxLength = .Fields("说明").DefinedSize
        
        Me.cbxSubClass.ListIndex = 0
        
        If .RecordCount > 0 Then
            Me.txt编号.Text = "" & !编号
            Me.txt名称.Text = "" & !名称
            Me.txt说明.Text = "" & !说明
            
            For i = 0 To cbxSubClass.ListCount - 1
                If Me.cbxSubClass.List(i) = NVL(!子类) Then
                    Me.cbxSubClass.ListIndex = i
                    Exit For
                End If
            Next i
            
            Select Case NVL(!保留, 0)
            Case 2
                chk报告.Value = vbChecked: optEditType(1).Value = True
            Case 0
                Select Case NVL(!通用, 0)
                Case 2
                    Me.chk报告.Value = vbChecked: Me.opt报告(1).Value = True
                Case 1
                    Me.chk报告.Value = vbChecked: Me.opt报告(0).Value = True
                Case Else
                    Me.chk报告.Value = vbUnchecked
                End Select
            End Select
            optApply(Val(!格式 & "")).Value = True
        End If
    End With
    
    gstrSQL = "Select 项目, Nvl(必填, 0) As 必填,nvl(只读,0) as 只读, Nvl(要素id, 0) As 要素, 内容 From 病历单据附项 Where 文件id = [1] Order By 排列"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngBillId)
    
    Err = 0: On Error GoTo 0
    With Me.vfg附项
        .Redraw = flexRDNone
        .Clear
        
        Set .DataSource = rsTemp
        
        .ColComboList(mCol.要素id) = mstrCombos
        .ColWidth(mCol.项目) = 1500: .ColWidth(mCol.必填) = 450: .ColWidth(mCol.只读) = 450: .ColWidth(mCol.要素id) = 1000
        
        For lngCount = 0 To .Cols - 1
            .FixedAlignment(lngCount) = flexAlignCenterCenter
        Next
        
        For lngCount = .FixedRows To .Rows - 1
            .Cell(flexcpChecked, lngCount, mCol.必填) = IIf(.TextMatrix(lngCount, mCol.必填) = "0", flexUnchecked, flexChecked)
            .Cell(flexcpChecked, lngCount, mCol.只读) = IIf(.TextMatrix(lngCount, mCol.只读) = "0", flexUnchecked, flexChecked)
            
            .TextMatrix(lngCount, mCol.必填) = ""
            .TextMatrix(lngCount, mCol.只读) = ""
        Next
        If .Rows > .FixedRows Then .Row = .FixedRows
        .Redraw = flexRDDirect
    End With
    LoadApplyFile
    zlRefresh = True: Exit Function

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlRefresh = False: Exit Function
End Function

Public Function zlEditStart(blnAdd As Boolean, lngBillId As Long) As Boolean
    '功能：开始项目编辑
    '参数： blnAdd-是否增加，否则为修改
    '       lngBillID-增加的参照项目，或者指定编辑的项目
    Dim rsTemp As New ADODB.Recordset
    
    If blnAdd Then
        Err = 0: On Error GoTo errHand
        gstrSQL = "Select Nvl(Max(To_Number(编号)), 0) As 编号, Nvl(Max(Length(编号)), 0) As 长度 From 病历文件列表 Where 种类 = 7"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取编号")
        
        If rsTemp!长度 <> 0 And rsTemp!长度 <= Me.txt编号.MaxLength Then
            Me.txt编号.Text = Format(Val(rsTemp!编号) + 1, String(rsTemp!长度, "0"))
        Else
            Me.txt编号.Text = Format(Val(rsTemp!编号) + 1, String(Me.txt编号.MaxLength, "0"))
        End If
        Me.txt名称.Text = "": Me.txt说明.Text = ""
    Else
        If optEditType(1).Value Then
            optEditType(0).Enabled = False: opt报告(0).Enabled = False: opt报告(1).Enabled = False
        Else
            optEditType(1).Enabled = False
        End If
    End If

    Me.Tag = IIf(blnAdd, "增加", "修改")
    Me.picEdit.Enabled = True: Call Form_Resize
    Me.txt编号.SetFocus
    zlEditStart = True: Exit Function

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlEditStart = False: Exit Function
End Function

Public Sub zlEditCancel()
    '功能：放弃正在进行的编辑
    Me.Tag = ""
    Me.picEdit.Enabled = False: Call Form_Resize
    Call Me.zlRefresh(mlngBillID)
End Sub

Public Function zlEditSave() As Long
    '功能：保存正在进行的编辑,并返回正在编辑项目id,保存失败返回0
Dim lngNewId As Long, strLists As String
Dim strSQL As String, rsTmp As ADODB.Recordset
Dim strRPTPass1 As String, strRPTPass2 As String
Dim lngCount As Long
    Dim i As Long, blnTrans As Boolean
    Static objRpt As clsReport
    
    '一般特性检查
    If Trim(Me.txt编号.Text) = "" Then
        MsgBox "请输入编号！", vbInformation, gstrSysName
        Me.txt编号.SetFocus: zlEditSave = 0: Exit Function
    End If
    
    If Val(Me.txt编号.Text) > Val(String(Me.txt编号.MaxLength, "9")) Then
        MsgBox "编号太大！", vbInformation, gstrSysName
        Me.txt编号.SetFocus: zlEditSave = 0: Exit Function
    End If
    
    If Me.Tag = "增加" Then
        strSQL = "Select ID, 编号, 名称, 说明, 通用 From 病历文件列表 Where 种类 = 7 And 编号 = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Format(Trim(Me.txt编号.Text), String(Me.txt编号.MaxLength, "0")))
        If rsTmp.RecordCount > 0 Then
            MsgBox "编号重复！", vbInformation, gstrSysName
            Me.txt编号.SetFocus: zlEditSave = 0: Exit Function
        End If
    End If
    
    If Trim(Me.txt名称.Text) = "" Then
        MsgBox "请输入名称！", vbInformation, gstrSysName
        Me.txt名称.SetFocus: zlEditSave = 0: Exit Function
    End If
    
    If LenB(StrConv(Trim(Me.txt名称.Text), vbFromUnicode)) > Me.txt名称.MaxLength Then
        MsgBox "名称超长（最多" & Me.txt名称.MaxLength & "个字符或等长汉字）！", vbInformation, gstrSysName
        Me.txt名称.SetFocus: zlEditSave = 0: Exit Function
    End If
    
    If LenB(StrConv(Trim(Me.txt说明.Text), vbFromUnicode)) > Me.txt说明.MaxLength Then
        MsgBox "说明超长（最多" & Me.txt说明.MaxLength & "个字符或等长汉字）！", vbInformation, gstrSysName
        Me.txt说明.SetFocus: zlEditSave = 0: Exit Function
    End If
    
    With Me.vfg附项
        strLists = ""
        If .Col = mCol.要素id Then .Select .Row, .Col + 1
        For lngCount = .FixedRows To .Rows - 1
            strLists = strLists & vbLf & .TextMatrix(lngCount, mCol.项目)
            
            If .Cell(flexcpChecked, lngCount, mCol.必填) = flexChecked Then
                strLists = strLists & vbTab & "1"
            Else
                strLists = strLists & vbTab & "0"
            End If
            
            If .Cell(flexcpChecked, lngCount, mCol.只读) = flexChecked Then
                strLists = strLists & vbTab & "1"
            Else
                strLists = strLists & vbTab & "0"
            End If
            
            
            If Val(.TextMatrix(lngCount, mCol.要素id)) = 0 Then
                strLists = strLists & vbTab
            Else
                strLists = strLists & vbTab & .TextMatrix(lngCount, mCol.要素id)
            End If
            
            strLists = strLists & vbTab & .TextMatrix(lngCount, mCol.内容)
        Next
        If strLists <> "" Then strLists = Mid(strLists, 2)
    End With
    
    '数据保存语句组织
    gstrSQL = "'" & Format(Trim(Me.txt编号.Text), String(Me.txt编号.MaxLength, "0")) & "','" & Trim(Me.txt名称.Text) & "','" & Trim(Me.txt说明.Text) & "'"
    If optEditType(1).Value Then   '保留
        gstrSQL = gstrSQL & ",2"
    Else
        gstrSQL = gstrSQL & ",0"
    End If
    
    If Me.chk报告.Value = vbUnchecked Then '通用
        gstrSQL = gstrSQL & ",0"
    ElseIf optEditType(0).Value Then
        If Me.opt报告(0).Value Then
            gstrSQL = gstrSQL & ",1"
        Else
            gstrSQL = gstrSQL & ",2"
        End If
    Else
        gstrSQL = gstrSQL & ",1"
    End If
    
    'ZLCISBILL00' || 编号_In || '-' || Form_In
    '11698 - 程序新增的报表，设计时报错。原因，报表密码未生成。
    If objRpt Is Nothing Then
        Set objRpt = New clsReport
        Call objRpt.InitOracle(gcnOracle)
    End If
    
    strRPTPass1 = objRpt.GenReportPass("ZLCISBILL00" & Format(Trim(Me.txt编号.Text), String(Me.txt编号.MaxLength, "0")) & "-1", Trim(Me.txt名称.Text))
    strRPTPass2 = objRpt.GenReportPass("ZLCISBILL00" & Format(Trim(Me.txt编号.Text), String(Me.txt编号.MaxLength, "0")) & "-2", Trim(Me.txt名称.Text))
    
    If Me.Tag = "增加" Then
        lngNewId = zlDatabase.GetNextId("病历文件列表")
        gstrSQL = "Zl_诊疗单据目录_Edit(1," & lngNewId & "," & gstrSQL & ",'" & strLists & "','" & Replace(strRPTPass1, "'", "''") & "','" & Replace(strRPTPass2, "'", "''") & "','" & cbxSubClass.Text & "'," & IIf(optApply(1).Value, 1, 0) & ")"
    Else
        lngNewId = mlngBillID
        gstrSQL = "Zl_诊疗单据目录_Edit(2," & lngNewId & "," & gstrSQL & ",'" & strLists & "','" & Replace(strRPTPass1, "'", "''") & "','" & Replace(strRPTPass2, "'", "''") & "','" & cbxSubClass.Text & "'," & IIf(optApply(1).Value, 1, 0) & ")"
    End If
    Err = 0: On Error GoTo errHand
    Screen.MousePointer = 11
    gcnOracle.BeginTrans: blnTrans = True
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存诊疗单据")
    
    For i = LBound(arrSQL) To UBound(arrSQL)
        If arrSQL(i) <> "" Then Call zlDatabase.ExecuteProcedure(arrSQL(i), Me.Caption)
    Next
    For i = LBound(arrSQLFile) To UBound(arrSQLFile)
        If arrSQLFile(i) <> "" Then Call zlDatabase.ExecuteProcedure(arrSQLFile(i), Me.Caption)
    Next

    gcnOracle.CommitTrans: blnTrans = False
    Screen.MousePointer = 0
    
    If Me.Tag = "增加" Then mlngBillID = lngNewId
    Me.Tag = ""
    Me.picEdit.Enabled = False: Call Form_Resize
    zlEditSave = mlngBillID: Exit Function
    
errHand:
    If blnTrans Then gcnOracle.RollbackTrans
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlEditSave = 0: Exit Function
End Function

'--------------------------------------------
'以下为窗体控件响应事件
'--------------------------------------------
Private Sub chk报告_Click()
    If Me.Tag <> "修改" Then
        If Me.chk报告.Value = vbUnchecked Then
            optEditType(0).Enabled = False: optEditType(0).Value = False
            optEditType(1).Enabled = False: optEditType(1).Value = False
            Me.opt报告(0).Enabled = False: Me.opt报告(0).Value = False
            Me.opt报告(1).Enabled = False: Me.opt报告(1).Value = False
        Else
            optEditType(0).Value = True: optEditType(0).Enabled = True
            optEditType(1).Enabled = True
            Me.opt报告(0).Value = True: Me.opt报告(0).Enabled = True
            Me.opt报告(1).Enabled = True
        End If
    Else
        If optEditType(0).Value Then
            optEditType(0).Enabled = chk报告.Value = 1: opt报告(0).Enabled = chk报告.Value = 1: opt报告(1).Enabled = chk报告.Value = 1
        Else
            optEditType(0).Enabled = False: optEditType(1).Enabled = False
            opt报告(0).Enabled = False: opt报告(1).Enabled = False
        End If
    End If
End Sub

Private Sub cmdEdit_Click(Index As Integer)
Dim strCell As String
Dim lngCount As Long
    With Me.vfg附项
        Select Case Index
        Case 0
            .Rows = .Rows + 1: .Row = .Rows - 1
            
            .Cell(flexcpChecked, .Row, mCol.必填) = flexUnchecked
            .Cell(flexcpChecked, .Row, mCol.只读) = flexUnchecked
            
            If .RowIsVisible(.Row) = False Then .TopRow = .Row
        Case 1
            If .Row < .FixedRows Then Exit Sub
            .RemoveItem .Row
        Case 2
            If .Row < .FixedRows + 1 Then Exit Sub
            For lngCount = 0 To .Cols - 1
                strCell = .TextMatrix(.Row - 1, lngCount)
                .TextMatrix(.Row - 1, lngCount) = .TextMatrix(.Row, lngCount)
                .TextMatrix(.Row, lngCount) = strCell
            Next
            .Row = .Row - 1
        Case 3
            If .Row < .FixedRows Then Exit Sub
            If .Row >= .Rows - 1 Then Exit Sub
            For lngCount = 0 To .Cols - 1
                strCell = .TextMatrix(.Row + 1, lngCount)
                .TextMatrix(.Row + 1, lngCount) = .TextMatrix(.Row, lngCount)
                .TextMatrix(.Row, lngCount) = strCell
            Next
            .Row = .Row + 1
        End Select
        If .Visible And .Editable Then .SetFocus
    End With
End Sub

Private Sub Form_Load()
    Dim strLists As String
    
    '载入单据所属分类
    Call LoadSubClass
    With vsFile
        .ColWidth(0) = 1500: .ColWidth(1) = 255
        .TextMatrix(0, 0) = "类型"
        .TextMatrix(0, 2) = "文件名"
        .TextMatrix(1, 0) = "xsl格式文件"
        .TextMatrix(2, 0) = "xml数据文件"
        .TextMatrix(3, 0) = "html显示文件"
        .Cell(flexcpAlignment, 0, 0, 0, 2) = flexAlignCenterCenter
        .Cell(flexcpAlignment, 0, 0, 3, 0) = flexAlignCenterCenter
    End With
    
    
    mlngBillID = 0: Me.picEdit.BackColor = Me.BackColor: picEditType.BackColor = Me.BackColor
    Err = 0: On Error GoTo errHand
    gstrSQL = "Select I.ID, I.中文名,decode(I.表示法,4,I.数值域,NULL) as 数值域" & vbNewLine & _
            "From 诊治所见项目 I, 诊治所见分类 K" & vbNewLine & _
            "Where I.分类id = K.ID And (K.性质 = 1 And K.编码 = '06' And" & vbNewLine & _
            "      I.中文名 Not In ('门诊诊断', '一次住院诊断', '二次住院诊断', '上次住院诊断')  Or k.性质 = 6) " & vbNewLine & _
            "Order By I.编码"
    strLists = ""
    
    Call zlDatabase.OpenRecordset(mrsItems, gstrSQL, "提取诊治所见项目")
    
    Do While Not mrsItems.EOF
        strLists = strLists & "|#" & mrsItems!ID & ";" & mrsItems!中文名
        mrsItems.MoveNext
    Loop
    
    mstrCombos = "#0; " & strLists
    
    Exit Sub

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub LoadSubClass()
'载入诊疗项目类别
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
    strSQL = "select 名称 from 诊疗项目类别 order by 名称"
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    cbxSubClass.Clear
    
    If rsData.RecordCount <= 0 Then Exit Sub
    
    cbxSubClass.AddItem ""
    While Not rsData.EOF
        Call cbxSubClass.AddItem(NVL(rsData!名称))
        
        Call rsData.MoveNext
    Wend
    
End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    With Me.vfg附项
        If Me.picEdit.Enabled = True Then
            Me.picEdit.BackColor = RGB(230, 230, 230)
            .FocusRect = flexFocusHeavy: .Editable = flexEDKbdMouse
            .Height = Me.cmdEdit(0).Top - .Top - Screen.TwipsPerPixelY
        Else
            Me.picEdit.BackColor = Me.BackColor
            .FocusRect = flexFocusNone: .Editable = flexEDNone
            .Height = Me.cmdEdit(0).Top + Me.cmdEdit(0).Height - .Top
        End If
    End With
    Me.chk报告.BackColor = Me.picEdit.BackColor
    Me.opt报告(0).BackColor = Me.picEdit.BackColor
    Me.opt报告(1).BackColor = Me.picEdit.BackColor
    Me.picEditType.BackColor = Me.picEdit.BackColor
    Me.optEditType(0).BackColor = Me.picEdit.BackColor
    Me.optEditType(1).BackColor = Me.picEdit.BackColor
    Me.Label1.BackColor = Me.picEdit.BackColor
    Me.picApplyType.BackColor = Me.picEdit.BackColor
    Me.optApply(0).BackColor = Me.picEdit.BackColor
    Me.optApply(1).BackColor = Me.picEdit.BackColor
    Me.picApply.BackColor = Me.picEdit.BackColor
    lblApplysMark.BackColor = Me.picEdit.BackColor
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mrsItems = Nothing
End Sub

Private Sub optApply_Click(Index As Integer)
    If Index = 1 Then
        picApply.Visible = True
        vfg附项.Visible = False
    Else
        picApply.Visible = False
        vfg附项.Visible = True
    End If
End Sub

Private Sub optEditType_Click(Index As Integer)
    If Index = 0 Then
        opt报告(0).Value = True: opt报告(0).Enabled = True
        opt报告(1).Enabled = True
    Else
        opt报告(0).Value = False: opt报告(1).Value = False
        opt报告(0).Enabled = False: opt报告(1).Enabled = False
    End If
End Sub

Private Sub txt编号_GotFocus()
    Me.txt编号.SelStart = 0: Me.txt编号.SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt编号_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt名称_GotFocus()
    Me.txt名称.SelStart = 0: Me.txt名称.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt名称_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr("%_'""", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt说明_GotFocus()
    Me.txt说明.SelStart = 0: Me.txt说明.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt说明_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0: Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr("%_'""", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub vfg附项_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col <> mCol.项目 Then Exit Sub
    Me.vfg附项.TextMatrix(Row, Col) = Replace(Me.vfg附项.TextMatrix(Row, Col), """", "")
    Me.vfg附项.TextMatrix(Row, Col) = Replace(Me.vfg附项.TextMatrix(Row, Col), "'", "")
    Me.vfg附项.TextMatrix(Row, Col) = Replace(Me.vfg附项.TextMatrix(Row, Col), vbTab, " ")
    Me.vfg附项.TextMatrix(Row, Col) = Replace(Me.vfg附项.TextMatrix(Row, Col), vbLf, " ")
End Sub

Private Sub vfg附项_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If Val(vfg附项.TextMatrix(NewRow, mCol.要素id)) <> 0 Then
        mrsItems.Filter = "ID=" & Val(vfg附项.TextMatrix(NewRow, mCol.要素id))
        If mrsItems.RecordCount > 0 Then
            mrsItems.MoveFirst
            vfg附项.ColComboList(mCol.内容) = " |" & Replace(mrsItems!数值域 & "", ";", "|")
        Else
           vfg附项.ColComboList(mCol.内容) = ""
        End If
    Else
        vfg附项.ColComboList(mCol.内容) = ""
    End If
End Sub

Private Sub vfg附项_DblClick()
    With Me.vfg附项
        If .Row < .FixedRows Then Exit Sub
        
        If .Col = mCol.必填 Then
            If .Cell(flexcpChecked, .Row, mCol.必填) = flexChecked Then
                .Cell(flexcpChecked, .Row, mCol.必填) = flexUnchecked
            Else
                .Cell(flexcpChecked, .Row, mCol.必填) = flexChecked
            End If
        End If
        
        If .Col = mCol.只读 Then
            If .Cell(flexcpChecked, .Row, mCol.只读) = flexChecked Then
                .Cell(flexcpChecked, .Row, mCol.只读) = flexUnchecked
            Else
                .Cell(flexcpChecked, .Row, mCol.只读) = flexChecked
            End If
        End If
        
    End With
End Sub

Private Sub vfg附项_KeyPress(KeyAscii As Integer)
    Call vfg附项_DblClick
End Sub

Private Sub vfg附项_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
    If Chr(KeyAscii) = """" Then KeyAscii = 0
End Sub

Private Sub vfg附项_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = mCol.必填 Or Col = mCol.只读 Then Cancel = True
End Sub

Private Sub vsFile_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    vsFile.Editable = flexEDKbdMouse
    vsFile.ColComboList(2) = "..."
End Sub

Private Function LoadApplyFile() As Boolean
'功能：显示临床路径文件内容和患者临床路径
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long

    On Error GoTo ErrH
    
    strSQL = "Select 文件名, 类别 From 自定义申请单文件 Where 文件id = [1] Order By 类别"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngBillID)
    With vsFile
        .TextMatrix(1, 2) = ""
        .TextMatrix(2, 2) = ""
        .TextMatrix(3, 2) = ""
        Set .Cell(flexcpPicture, 1, 1) = Nothing
        Set .Cell(flexcpPicture, 2, 1) = Nothing
        Set .Cell(flexcpPicture, 3, 1) = Nothing
        For i = 1 To rsTmp.RecordCount
            .TextMatrix(Val(rsTmp!类别 & ""), 2) = rsTmp!文件名
            Set .Cell(flexcpPicture, Val(rsTmp!类别 & ""), 1) = zlCommFun.GetFileIcon(rsTmp!文件名, True, App.hInstance)
            .Cell(flexcpPictureAlignment, Val(rsTmp!类别 & ""), 1) = flexAlignCenterCenter
            rsTmp.MoveNext
        Next
    End With

    LoadApplyFile = True
    Exit Function
ErrH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub vsFile_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim arrTmp() As String
    Dim strFile As String
    Dim strFileName As String
    Dim strSQL As String
    Dim StrText As String
    Dim i As Long
    Dim stmTmp As Stream
    
	If mblndb = True Then mblndb = False: Exit Sub
    If Row = 1 Then
        cdgFile.DialogTitle = "选择要添加的xsl格式文件"
        cdgFile.Filter = "html显示文件(*.xsl)|*.xsl"
    ElseIf Row = 2 Then
        cdgFile.DialogTitle = "选择要添加的xml数据文件"
        cdgFile.Filter = "html显示文件(*.xml)|*.xml"
    ElseIf Row = 3 Then
        cdgFile.DialogTitle = "选择要添加的html显示文件"
        cdgFile.Filter = "html显示文件(*.html)|*.html"
    End If
    cdgFile.flags = &H80000 Or &H4 Or &H1000 Or &H200000 Or &H800
    cdgFile.InitDir = GetSetting("ZLSOFT", "公共模块\" & App.ProductName, "自定义申请单选择目录")
    cdgFile.CancelError = True
    On Error Resume Next
    cdgFile.ShowOpen
    If Err.Number <> 0 Then
        Err.Clear: Exit Sub
    End If
    On Error GoTo ErrH
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName, "自定义申请单选择目录", mobjFile.GetFile(cdgFile.Filename).ParentFolder.Path
    strFile = cdgFile.Filename '包含路径
    strFileName = mobjFile.GetFile(cdgFile.Filename).Name
    
    '检查文件大小不超过3M
    If mobjFile.GetFile(strFile).Size / 1024 / 1024 > 3 Then
        MsgBox "文件尺寸太大(超过3M)，请对文件进行适当的整理后再添加。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    On Error Resume Next
    Set stmTmp = New Stream
    stmTmp.Open
    stmTmp.Charset = "UTF-8"
    stmTmp.Position = 0
    stmTmp.LoadFromFile strFile
    stmTmp.Position = 0
    
    StrText = stmTmp.ReadText
    stmTmp.Close

    If Err.Number <> 0 Then
        MsgBox "文件添加失败,请检查文件格式及文件内容是否正确！", vbExclamation, gstrSysName
        Screen.MousePointer = 0: Err.Clear: Exit Sub
    End If
    On Error GoTo ErrH

    ReDim arrTmp(0)
    strSQL = "Zl_自定义申请单文件_Edit(1," & mlngBillID & "," & Row & ",'" & strFileName & "')"
    If Not Sys.GetlobSql(glngSys, 24, mlngBillID & "," & Row, Replace(StrText, "'", "''"), arrTmp(), 1) Then
        MsgBox "文件添加失败！", vbExclamation, gstrSysName
        Screen.MousePointer = 0: Exit Sub
    End If
    If arrSQL(UBound(arrSQL)) <> "" Then ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = strSQL
    For i = LBound(arrTmp) To UBound(arrTmp)
        If arrSQLFile(UBound(arrSQLFile)) <> "" Then ReDim Preserve arrSQLFile(UBound(arrSQLFile) + 1)
        arrSQLFile(UBound(arrSQLFile)) = arrTmp(i)
    Next

    vsFile.TextMatrix(Row, 2) = strFileName
    Set vsFile.Cell(flexcpPicture, Row, 1) = zlCommFun.GetFileIcon(strFileName, True, App.hInstance)
    vsFile.Cell(flexcpPictureAlignment, Row, 1) = flexAlignCenterCenter
    Exit Sub
ErrH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsFile_DblClick()
    Dim strFile As String
    Dim lngRetu As Long, strInfo As String
    Dim StrText As String
    Dim stmTmp As Stream
    
On Error Resume Next
    Screen.MousePointer = 11
    mblndb = True
    strFile = mobjFile.GetSpecialFolder(TemporaryFolder) & "\" & vsFile.TextMatrix(vsFile.Row, 2)
    If mobjFile.FileExists(strFile) Then mobjFile.DeleteFile strFile, True
    
    StrText = Sys.Readlob(glngSys, 24, mlngBillID & "," & vsFile.Row, strFile, 1)
    
    If Not mobjFile.FileExists(strFile) Then Call mobjFile.CreateTextFile(strFile)
    Set stmTmp = New Stream
    stmTmp.Open
    stmTmp.Charset = "UTF-8"
    stmTmp.WriteText StrText
    stmTmp.SaveToFile strFile, adSaveCreateOverWrite
    stmTmp.Close
    
    If Not mobjFile.FileExists(strFile) Then
        MsgBox "文件内容读取失败！", vbInformation, gstrSysName:
        Screen.MousePointer = 0: Exit Sub
    End If
    
    lngRetu = ShellExecute(Me.hWnd, "open", strFile, "", "", SW_SHOWNORMAL)
    If lngRetu <= 32 Then
        Select Case lngRetu
        Case 2: strInfo = "错误的关联"
        Case 29: strInfo = "关联失败"
        Case 30: strInfo = "关联应用程式忙碌中..."
        Case 31: strInfo = "没有关联任何应用程式"
        Case Else: strInfo = "无法识别的错误"
        End Select
        MsgBox "文件打开时出错：" & vbCrLf & vbCrLf & vbTab & strInfo, vbExclamation, gstrSysName
    End If
    'If mobjFile.FileExists(strFile) Then mobjFile.DeleteFile strFile, True
    
    Screen.MousePointer = 0
End Sub
