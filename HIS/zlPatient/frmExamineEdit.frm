VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmExamineEdit 
   BackColor       =   &H00808080&
   Caption         =   "病人费用审批编辑"
   ClientHeight    =   7335
   ClientLeft      =   0
   ClientTop       =   375
   ClientWidth     =   10665
   Icon            =   "frmExamineEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7335
   ScaleWidth      =   10665
   StartUpPosition =   1  '所有者中心
   Begin VSFlex8Ctl.VSFlexGrid vsExist 
      Height          =   2370
      Left            =   2520
      TabIndex        =   0
      Top             =   480
      Width           =   7065
      _cx             =   12462
      _cy             =   4180
      Appearance      =   1
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
      BackColorSel    =   16574424
      ForeColorSel    =   0
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmExamineEdit.frx":1601A
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   5
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
   Begin VB.CommandButton cmd调用 
      Caption         =   "调用模板(&T)"
      Height          =   350
      Left            =   6600
      TabIndex        =   13
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "保存为模板(&S)"
      Height          =   350
      Left            =   8040
      TabIndex        =   12
      Top             =   120
      Width           =   1455
   End
   Begin MSComctlLib.TabStrip tabClass 
      Height          =   345
      Left            =   2520
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   7035
      _ExtentX        =   12409
      _ExtentY        =   609
      TabWidthStyle   =   1
      TabFixedWidth   =   2290
      TabFixedHeight  =   526
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "全部(&0)"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "西成药"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "中成药"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "中草药"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "治疗"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDropMode     =   1
   End
   Begin VB.PictureBox picLineX 
      BorderStyle     =   0  'None
      Height          =   45
      Left            =   0
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   8415
      TabIndex        =   7
      Top             =   2760
      Width           =   8415
   End
   Begin VB.PictureBox picLineY 
      BorderStyle     =   0  'None
      Height          =   3375
      Left            =   4440
      MousePointer    =   9  'Size W E
      ScaleHeight     =   3375
      ScaleWidth      =   45
      TabIndex        =   6
      Top             =   2880
      Width           =   45
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   400
      Left            =   2550
      ScaleHeight     =   405
      ScaleWidth      =   5775
      TabIndex        =   2
      Top             =   2955
      Width           =   5775
      Begin VB.CommandButton cmdDelete 
         Caption         =   "删除(&D)"
         Height          =   350
         Left            =   0
         TabIndex        =   5
         Top             =   30
         Width           =   1100
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "增加(&A)"
         Height          =   350
         Left            =   1125
         TabIndex        =   4
         Top             =   30
         Width           =   1100
      End
      Begin VB.TextBox txtFind 
         Height          =   300
         Left            =   2775
         TabIndex        =   3
         Top             =   55
         Width           =   2955
      End
      Begin VB.Image ImgFind 
         Height          =   240
         Left            =   2475
         Picture         =   "frmExamineEdit.frx":1615C
         Top             =   85
         Width           =   240
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsList 
      Height          =   2865
      Left            =   2520
      TabIndex        =   8
      Top             =   3375
      Width           =   7065
      _cx             =   12462
      _cy             =   5054
      Appearance      =   1
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
      BackColorSel    =   16574424
      ForeColorSel    =   0
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmExamineEdit.frx":164E6
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   5
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
   Begin MSComctlLib.TreeView tvwMain_S 
      Height          =   3285
      Left            =   0
      TabIndex        =   9
      Top             =   2955
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   5794
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      HotTracking     =   -1  'True
      ImageList       =   "ils16"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExamineEdit.frx":165BF
            Key             =   "RootS"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExamineEdit.frx":16719
            Key             =   "Exp"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExamineEdit.frx":16873
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExamineEdit.frx":16CC5
            Key             =   "RootR"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExamineEdit.frx":17117
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExamineEdit.frx":1756F
            Key             =   "ItemNo"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExamineEdit.frx":179C3
            Key             =   "Write"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExamineEdit.frx":17E17
            Key             =   "Read"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExamineEdit.frx":1826B
            Key             =   "ItemR"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExamineEdit.frx":190BD
            Key             =   "ItemRNo"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   10
      Top             =   6975
      Width           =   10665
      _ExtentX        =   18812
      _ExtentY        =   635
      SimpleText      =   $"frmExamineEdit.frx":19F0F
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmExamineEdit.frx":19F56
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13732
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ListView lvw模板 
      Height          =   2895
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   5106
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ils16"
      SmallIcons      =   "ils16"
      ColHdrIcons     =   "ils16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lbl分类 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "审批项目分类"
      Height          =   180
      Left            =   0
      TabIndex        =   16
      Top             =   7320
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Label lbl项目 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "模板项目列表"
      Height          =   180
      Left            =   0
      TabIndex        =   15
      Top             =   7080
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Label lbl模板 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "审批项目模板列表"
      Height          =   180
      Left            =   0
      TabIndex        =   14
      Top             =   6840
      Visible         =   0   'False
      Width           =   1440
   End
End
Attribute VB_Name = "frmExamineEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum ColAdded
    选择 = 0: 使用限量: 类别: 编码: 名称: 规格: 产地: 单位: 说明: 审批人: 审批时间: ID
End Enum

Private Enum ColAdd
    选择 = 0: 类别: 编码: 名称: 规格: 产地: 单位: 说明: ID
End Enum

Private mrsExistItem As New ADODB.Recordset
Private mint简码 As Integer
Private mlng病人ID As Long, mlng主页ID As Long, mlng险类 As Long
Private mlngCount As Long
Private mblnDel As Boolean
Private mblnDelPrv As Boolean
Private mbln模板 As Boolean

Private Function Exist编码(lng编码 As Long, str名称 As String) As Boolean
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    strSQL = "select 1 from 审批项目模板 where 编码=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng编码)
    If Not rsTemp.EOF Then Exist编码 = True: Exit Function
    
    strSQL = "select 1 from 审批项目模板 where 名称=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str名称)
    If Not rsTemp.EOF Then Exist编码 = True: Exit Function
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub ReadExistsTemplet()
'读取已存在的模板
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim lvwItem As ListItem
    
On Error GoTo errHandle
    
    If mbln模板 Then
        strSQL = "Select distinct(编码) 编码,名称 From 审批项目模板 Order By 编码"
    Else
        strSQL = "Select distinct(编码) 编码,名称 From 审批项目模板 A,保险支付项目 B Where A.项目ID = B.收费细目ID And B.要求审批 = 1 And B.险类 = [1] Order By 编码"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng险类)
    lvw模板.ListItems.Clear
    
    If mbln模板 = True Then
        Set lvwItem = lvw模板.ListItems.Add(, "_ADD", "新增审批项目模板", "Write", "Write")
    End If
    '问题30028 by lesfeng 2010-06-01 解决名称为空情况
    While Not rsTemp.EOF
        Set lvwItem = lvw模板.ListItems.Add(, "_" & rsTemp!编码, IIf(IsNull(rsTemp!名称), "", rsTemp!名称), "Item", "Item")
        rsTemp.MoveNext
    Wend
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub View状态()
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle
    
    If mbln模板 = False Then
        strSQL = "Select 住院号, 姓名 From 病人信息 Where 病人id = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng病人ID)
        
        stbThis.Panels(2).Text = "住院号:" & rsTemp!住院号 & " 姓名:" & rsTemp!姓名 & "   该病人总共设定了" & mlngCount & "条审批记录！"
    Else
        If lvw模板.SelectedItem Is Nothing Then
            stbThis.Panels(2).Text = "新增模板请选择模板列表中的[新增审核项目模板],选择审核项目后点击[增加]"
        ElseIf UCase(Mid(lvw模板.SelectedItem.Key, 2)) = "ADD" Then
            stbThis.Panels(2).Text = "新增模板,选择审核项目后点击[增加]"
        Else
            stbThis.Panels(2).Text = "模板:" & lvw模板.SelectedItem.Text & "  该模板总共设定了" & mlngCount & "条审批记录！"
        End If
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdAdd_Click()
    Dim lngRow As Long, lngCount As Long
    Dim strTemp As String
    Dim blnTemp As Boolean, blnTrans As Boolean
    Dim lng编码 As Long
    Dim str名称 As String
        
    If mblnDel = True Then Unload Me: Exit Sub
    
    For lngRow = 1 To vsList.Rows - 1
        If vsList.Cell(flexcpChecked, lngRow, ColAdd.选择) = flexChecked Then
            blnTemp = True
        End If
    Next lngRow
    
    If blnTemp = False Then
        MsgBox "请选择要增加的审批项目!", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If mbln模板 = True Then
        If UCase(Mid(lvw模板.SelectedItem.Key, 2)) = "ADD" Then
            str名称 = frmTempletEdit.EditTemplet(Me)
            If str名称 = "" Then Exit Sub
            lng编码 = Mid(str名称, 1, InStr(str名称, ",") - 1)
            str名称 = Mid(str名称, InStr(str名称, ",") + 1)
            If Exist编码(lng编码, str名称) = True Then
                MsgBox "该编码已经存在不能新增!", vbInformation, gstrSysName
                Exit Sub
            End If
        Else
            lng编码 = Mid(lvw模板.SelectedItem.Key, 2)
            str名称 = lvw模板.SelectedItem.Text
        End If
    End If
    
On Error GoTo errHandle
    gcnOracle.BeginTrans: blnTrans = True
    
    For lngRow = 1 To vsList.Rows - 1
        If vsList.Cell(flexcpChecked, lngRow, ColAdd.选择) = flexChecked Then
            strTemp = vsList.TextMatrix(lngRow, ColAdd.ID) & "," & strTemp
            lngCount = lngCount + 1
        End If
        If lngCount = 100 Then
            If mbln模板 = False Then
                gstrSQL = "Zl_病人审批项目_Insert(" & mlng病人ID & "," & mlng主页ID & ",'" & strTemp & "','" & gstrUserName & "')"
                Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            Else
                gstrSQL = "zl_审批项目模板_Insert(" & lng编码 & ",'" & str名称 & "','" & strTemp & "')"
                Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            End If
            lngCount = 0
            strTemp = ""
        End If
    Next lngRow
    
    If strTemp <> "" Then
        If mbln模板 = False Then
            gstrSQL = "Zl_病人审批项目_Insert(" & mlng病人ID & "," & mlng主页ID & ",'" & strTemp & "','" & UserInfo.姓名 & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        Else
            gstrSQL = "zl_审批项目模板_Insert(" & lng编码 & ",'" & str名称 & "','" & strTemp & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        End If
    End If
    gcnOracle.CommitTrans: blnTrans = False
    
    vsList.Redraw = flexRDNone
    For lngRow = 1 To vsList.Rows - 1
        If vsList.Cell(flexcpChecked, lngRow, ColAdd.选择) = flexChecked Then
            vsList.Cell(flexcpChecked, lngRow, ColAdd.选择) = 2
            vsList.RowHidden(lngRow) = True
        End If
    Next lngRow
    
    vsList.Redraw = flexRDDirect
    If mbln模板 = False Then
        Call ReadExistsItem(mlng病人ID, mlng主页ID)
    Else
        If UCase(Mid(lvw模板.SelectedItem.Key, 2)) = "ADD" Then
            lvw模板.ListItems.Add , "_" & lng编码, str名称, "Item", "Item"
            Set lvw模板.SelectedItem = lvw模板.ListItems.Item("_" & lng编码)
            Call ReadTempletItem(lng编码)
        Else
            Call ReadTempletItem(lng编码)
        End If
    End If
    
    Exit Sub
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdDelete_Click()
    Dim strTemp As String
    Dim lngRow As Long
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim blnTemp As Boolean, blnTrans As Boolean
    Dim lng编码 As Long
    Dim i As Integer
    Dim strMsg As String
    
    '判断是否选择了要删除的项目
    For lngRow = 1 To vsExist.Rows - 1
      If vsExist.Cell(flexcpChecked, lngRow, ColAdded.选择) = 1 Then
        blnTemp = True
        Exit For
      End If
    Next lngRow
    
    If blnTemp = True Then
        If MsgBox("确定删除该审批项目吗?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    Else
        MsgBox "请选择要删除的项目!", vbInformation, Me.Caption
        vsExist.SetFocus
        Exit Sub
    End If
    
    On Error GoTo errHandle
    Screen.MousePointer = 11
    
'不是删除状态
    If mbln模板 = False Then
        
        For lngRow = 1 To vsExist.Rows - 1
            If vsExist.Cell(flexcpChecked, lngRow, ColAdded.选择) = 1 Then
                strTemp = strTemp & "," & vsExist.TextMatrix(vsExist.Row, ColAdded.ID)
            End If
        Next lngRow
        'by lesfeng 2009-12-30 大表拆分  病人费用记录 --〉住院费用记录 这里只对住院
        'by lesfeng 2010-03-06 性能绑定
        strSQL = "Select A.收费细目id, B.名称 " & _
                 "From 住院费用记录 A, 收费项目目录 B " & _
                 "Where A.收费细目id = B.ID And InStr([1], ',' || A.收费细目id || ',') > 0" & _
                    " And A.病人id = [2] And A.主页id = [3]" & _
                    " And (B.站点=[4] Or B.站点 is Null)"
    
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, "," & strTemp & ",", mlng病人ID, mlng主页ID, gstrNodeNo)
        
        If rsTemp.RecordCount > 0 Then
           Do While Not rsTemp.EOF
              If i > 5 Then strMsg = strMsg & "...": Exit Do
              strMsg = strMsg & ",[" & rsTemp!名称 & "]"
              i = i + 1
           Loop
           strMsg = Mid(strMsg, 2)
           If MsgBox("该病人已经存在" & strMsg & Chr(13) & Chr(10) & "的费用信息,是否继续?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        End If
        
        gcnOracle.BeginTrans: blnTrans = True
        For lngRow = 1 To vsExist.Rows - 1
            If vsExist.Cell(flexcpChecked, lngRow, ColAdded.选择) = 1 Then
                gstrSQL = "Zl_病人审批项目_Delete(" & mlng病人ID & "," & mlng主页ID & "," & vsExist.TextMatrix(lngRow, ColAdded.ID) & ")"
                Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            End If
        Next lngRow
        gcnOracle.CommitTrans: blnTrans = False
    Else
        '增加模板状态
        lng编码 = Mid(lvw模板.SelectedItem.Key, 2)
        
        gcnOracle.BeginTrans: blnTrans = True
        For lngRow = 1 To vsExist.Rows - 1
            If vsExist.Cell(flexcpChecked, lngRow, ColAdded.选择) = 1 Then
                gstrSQL = "ZL_审批项目模板_DELETE(" & lng编码 & "," & vsExist.TextMatrix(lngRow, ColAdded.ID) & ")"
                Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            End If
        Next lngRow
        gcnOracle.CommitTrans: blnTrans = False
    End If
    
    If mbln模板 = False Then
        Call ReadExistsItem(mlng病人ID, mlng主页ID)
    Else
        Call ReadTempletItem(lng编码)
        If mrsExistItem.RecordCount = 0 Then
            lvw模板.ListItems.Remove lvw模板.SelectedItem.Key
        End If
    End If
    Screen.MousePointer = 0
    Exit Sub
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdSave_Click()
    Dim lngRow As Long, lngCount As Long, blnTrans As Boolean
    Dim strTemp As String
    Dim lng编码 As Long
    Dim str名称 As String
    Dim strTabKey As String

    
    If vsExist.Rows < 1 Then
        MsgBox "没有保存为模板的审批项目!", vbInformation, gstrSysName
        Exit Sub
    End If
    
    str名称 = frmTempletEdit.EditTemplet(Me)
    If str名称 = "" Then Exit Sub
    lng编码 = Mid(str名称, 1, InStr(str名称, ",") - 1)
    str名称 = Mid(str名称, InStr(str名称, ",") + 1)
    If Exist编码(lng编码, str名称) = True Then
        MsgBox "该编码已经存在不能新增!", vbInformation, gstrSysName
        Exit Sub
    End If
    
On Error GoTo errHandle
    gcnOracle.BeginTrans: blnTrans = True
    strTabKey = tabClass.SelectedItem.Key
    mrsExistItem.Filter = 0
    
    While Not mrsExistItem.EOF
        
        strTemp = mrsExistItem!ID & "," & strTemp
        lngCount = lngCount + 1
        
        If lngCount = 100 Then
            gstrSQL = "zl_审批项目模板_Insert(" & lng编码 & ",'" & str名称 & "','" & strTemp & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            lngCount = 0
            strTemp = ""
        End If
        
        mrsExistItem.MoveNext
    Wend
    
    If strTemp <> "" Then
        gstrSQL = "zl_审批项目模板_Insert(" & lng编码 & ",'" & str名称 & "','" & strTemp & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    End If
    '问题30020 by lesfeng 2010-06-01 解决名称为空情况
    If Not mrsExistItem.EOF Then
        If str名称 <> "" Then
            lvw模板.ListItems.Add , "_" & lng编码, str名称, "Item", "Item"
        End If
    End If
    gcnOracle.CommitTrans: blnTrans = False
    
    Call tabClass_Click
    Exit Sub
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmd调用_Click()
    If lvw模板.ListItems.Count = 0 Then
        MsgBox "没有设置适用于当前病人险类的模板,请先设置模板!", vbInformation, gstrSysName
        Exit Sub
    End If
    lvw模板.Visible = True
    lvw模板.Move cmd调用.Left + cmd调用.Width - lvw模板.Width, cmd调用.Top + cmd调用.Height, lvw模板.Width, 3000
    lvw模板.Height = lvw模板.ListItems.Item(1).Height * (lvw模板.ListItems.Count + 1)
    lvw模板.ZOrder
    lvw模板.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = 2 Then
        If mblnDel = False Then
            If Me.ActiveControl.Name = "vsList" Then
                If vsList.Rows > 1 Then
                    vsList.Editable = flexEDKbdMouse
                    If KeyCode = vbKeyR Then
                        vsList.Cell(flexcpChecked, 1, ColAdd.选择, vsList.Rows - 1, ColAdd.选择) = 2
                    ElseIf KeyCode = vbKeyA Then
                        vsList.Cell(flexcpChecked, 1, ColAdd.选择, vsList.Rows - 1, ColAdd.选择) = 1
                    End If
                    vsList.Editable = flexEDNone
                End If
            ElseIf Me.ActiveControl.Name = "vsExist" Then
                If vsExist.Rows > 1 Then
                    vsExist.Editable = flexEDKbdMouse
                    If KeyCode = vbKeyR Then
                        vsExist.Cell(flexcpChecked, 1, ColAdd.选择, vsExist.Rows - 1, ColAdd.选择) = 2
                    ElseIf KeyCode = vbKeyA Then
                        vsExist.Cell(flexcpChecked, 1, ColAdd.选择, vsExist.Rows - 1, ColAdd.选择) = 1
                    End If
                    vsExist.Editable = flexEDNone
                End If
                
            End If
        Else
            If vsExist.Rows > 1 Then
                vsExist.Editable = flexEDKbdMouse
                If KeyCode = vbKeyR Then
                    vsExist.Cell(flexcpChecked, 1, ColAdd.选择, vsExist.Rows - 1, ColAdd.选择) = 2
                ElseIf KeyCode = vbKeyA Then
                    vsExist.Cell(flexcpChecked, 1, ColAdd.选择, vsExist.Rows - 1, ColAdd.选择) = 1
                End If
                vsExist.Editable = flexEDNone
            End If
        End If
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    picLineX.Width = Me.Width
    If mbln模板 = False Then
        vsExist.Move 0, tabClass.Height - 20, Me.Width - 100, picLineX.Top - tabClass.Height + 50
        tabClass.Left = 0
        tabClass.Top = 30
        If mblnDel = False Then
            tvwMain_S.Move 0, picLineX.Top + picLineX.Height, picLineY.Left, Me.ScaleHeight - picLineX.Top - picLineX.Height - stbThis.Height
            picLineY.Top = tvwMain_S.Top
            picLineY.Height = tvwMain_S.Height
            pic.Move picLineY.Left + picLineY.Width, tvwMain_S.Top, Me.Width - picLineY.Left - picLineY.Width, pic.Height
            vsList.Move pic.Left, tvwMain_S.Top + pic.Height + 20, pic.Width - 100, tvwMain_S.Height - pic.Height - 20
            tvwMain_S.Visible = True
            txtFind.Width = Me.Width - pic.Left - txtFind.Left - 200
            If cmdSave.Visible = True Then
                cmdSave.Move Me.ScaleWidth - cmdSave.Width - 100, 30, cmdSave.Width, 300
                cmd调用.Move cmdSave.Left - cmd调用.Width - 30, 30, cmd调用.Width, 300

            Else
                cmd调用.Move Me.ScaleWidth - cmd调用.Width - 100, 30, cmd调用.Width, 300
            End If
            tabClass.Width = cmd调用.Left - 100
            lvw模板.ColumnHeaders.Item(1).Width = lvw模板.Width - 100
            cmdSave.ZOrder
            cmd调用.ZOrder
            vsExist.ZOrder
        Else
            picLineY.Visible = False
            picLineX.Visible = False
            tvwMain_S.Visible = False
            vsList.Visible = False
            txtFind.Visible = False
            ImgFind.Visible = False
            stbThis.Visible = False
            vsExist.Height = Me.ScaleHeight - IIf(tabClass.Visible = True, tabClass.Height, 0) - pic.Height
            pic.Top = vsExist.Height + vsExist.Top
            pic.Width = cmdDelete.Width + cmdAdd.Width
            pic.Left = Me.Width - pic.Width - 200
            vsExist.ZOrder
        End If
    Else
        picLineY.Height = Me.ScaleHeight - stbThis.Height
        lvw模板.Visible = True
        cmd调用.Visible = False
        cmdSave.Visible = False
        lbl模板.Left = 30
        lbl模板.Top = 30
        lvw模板.Move 0, lbl模板.Height + lbl模板.Top, picLineY.Left, picLineX.Top - lbl模板.Height - lbl模板.Top
        If tabClass.Visible = False Then
            lbl项目.Left = picLineY.Left + picLineY.Width + 30
            lbl项目.Top = 30
            vsExist.Move picLineY.Left + picLineY.Width, lbl项目.Height + lbl项目.Top, Me.ScaleWidth - picLineY.Left - picLineY.Width, picLineX.Top - lbl项目.Height - lbl项目.Top
        Else
            lbl项目.Left = picLineY.Left + picLineY.Width
            lbl项目.Top = 30
            tabClass.Top = lbl项目.Height + lbl项目.Top
            vsExist.Move picLineY.Left + picLineY.Width, tabClass.Top + tabClass.Height - 30, Me.ScaleWidth - picLineY.Left - picLineY.Width, picLineX.Top - tabClass.Height + 30 - lbl项目.Height - lbl项目.Top
            tabClass.Left = picLineY.Left + picLineY.Width
            tabClass.Width = vsExist.Width
        End If
        lbl分类.Left = 30
        lbl分类.Top = picLineX.Top + picLineX.Height + 30
        tvwMain_S.Move 0, lbl分类.Top + lbl分类.Height, picLineY.Left, Me.ScaleHeight - picLineX.Top - picLineX.Height - stbThis.Height - lbl分类.Height - 30
        picLineY.Top = Me.ScaleTop
        pic.Move picLineY.Left + picLineY.Width, tvwMain_S.Top, Me.Width - picLineY.Left - picLineY.Width, pic.Height
        vsList.Move pic.Left, tvwMain_S.Top + pic.Height + 20, pic.Width - 100, tvwMain_S.Height - pic.Height - 20
        tvwMain_S.Visible = True
        txtFind.Width = Me.Width - pic.Left - txtFind.Left - 200
        stbThis.Visible = True
        lvw模板.ColumnHeaders.Item(1).Width = lvw模板.Width - 100
        vsExist.ZOrder
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
    Set mrsExistItem = Nothing
'    mblnDel = False
'    mblnDelPrv = False
    mlngCount = 0
    
End Sub

Private Sub ImgFind_Click()
    Call txtFind_KeyPress(vbKeyReturn)
End Sub

Private Sub lvw模板_DblClick()
    Dim strSQL As String, blnTrans As Boolean
    Dim rsTemp As ADODB.Recordset
    Dim strTemp As String
    Dim lngCount As Long
    
On Error GoTo errHandle
   
    If mbln模板 = False Then
        If MsgBox("该病人的审批项目是否加载模版" & lvw模板.SelectedItem.Text & "里的内容?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then lvw模板.Visible = False: Exit Sub
        
        strSQL = "Select A.项目id From 审批项目模板 A ,保险支付项目 B Where A.项目id = B.收费细目ID And B.险类 = [4] And A.编码 = [1] And B.要求审批 = 1 " & _
                 "Minus " & _
                 " Select 项目ID From 病人审批项目 A Where A.病人id = [2] And 主页ID = [3]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(lvw模板.SelectedItem.Key, 2), mlng病人ID, mlng主页ID, mlng险类)
        
        gcnOracle.BeginTrans: blnTrans = True
        While Not rsTemp.EOF
            strTemp = rsTemp!项目id & "," & strTemp
            lngCount = lngCount + 1
            
            If lngCount = 100 Then
                gstrSQL = "Zl_病人审批项目_Insert(" & mlng病人ID & "," & mlng主页ID & ",'" & strTemp & "','" & UserInfo.姓名 & "')"
                Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                lngCount = 0
                strTemp = ""
            End If
            
            rsTemp.MoveNext
        Wend
        
        If strTemp <> "" Then
             gstrSQL = "Zl_病人审批项目_Insert(" & mlng病人ID & "," & mlng主页ID & ",'" & strTemp & "','" & UserInfo.姓名 & "')"
             Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        End If
        gcnOracle.CommitTrans: blnTrans = False
        
        Call ReadExistsItem(mlng病人ID, mlng主页ID)
        vsList.Tag = ""
        If Not tvwMain_S.SelectedItem Is Nothing Then Call tvwMain_S_NodeClick(tvwMain_S.SelectedItem)
        lvw模板.Visible = False
    End If
    Exit Sub
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub lvw模板_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If mbln模板 = True Then
        If UCase(Mid(lvw模板.SelectedItem.Key, 2)) <> "ADD" Then
            If lvw模板.Tag <> lvw模板.SelectedItem.Key Then
                Call ReadTempletItem(Mid(lvw模板.SelectedItem.Key, 2))
                vsList.Tag = ""
                If Not tvwMain_S.SelectedItem Is Nothing Then Call tvwMain_S_NodeClick(tvwMain_S.SelectedItem)
                lvw模板.Tag = lvw模板.SelectedItem.Key
            End If
        Else
            Set mrsExistItem = Nothing
            tabClass.Visible = False
            vsList.Tag = ""
            If Not tvwMain_S.SelectedItem Is Nothing Then Call tvwMain_S_NodeClick(tvwMain_S.SelectedItem)
            Call Form_Resize
            lvw模板.Tag = ""
            vsExist.Rows = 1
            View状态
        End If
    End If
End Sub

Private Sub lvw模板_KeyDown(KeyCode As Integer, Shift As Integer)
    If mbln模板 = False Then
        If KeyCode = vbKeyEscape Then
            lvw模板.Visible = False
        End If
    End If
End Sub

Private Sub lvw模板_KeyPress(KeyAscii As Integer)
    If mbln模板 = False Then
        If KeyAscii = vbKeyReturn Then
            Call lvw模板_DblClick
        End If
    End If
End Sub

Private Sub picLineX_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    
    With picLineX
        If .Top + Y < 1000 Then Exit Sub
        If .Top + Y > Me.ScaleHeight - 2000 Then Exit Sub
        .Top = .Top + Y
    End With
    Call Form_Resize
    Me.Refresh
End Sub

Private Sub picLineY_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    
    With picLineY
        If .Left + X < 2000 Then Exit Sub
        If .Left + X > Me.ScaleWidth - 3500 Then Exit Sub
        
        .Move .Left + X
    End With
    Call Form_Resize
    Me.Refresh
End Sub

Private Sub InitForm()
'初始化界面
    With vsExist
        .Rows = 2
        .Cols = 12
        .TextMatrix(0, ColAdded.选择) = ""
        .TextMatrix(0, ColAdded.使用限量) = "使用限量"
        .TextMatrix(0, ColAdded.类别) = "类别"
        .TextMatrix(0, ColAdded.名称) = "名称"
        .TextMatrix(0, ColAdded.编码) = "编码"
        .TextMatrix(0, ColAdded.规格) = "规格"
        .TextMatrix(0, ColAdded.产地) = "产地"
        .TextMatrix(0, ColAdded.单位) = "单位"
        .TextMatrix(0, ColAdded.说明) = "说明"
        .TextMatrix(0, ColAdded.审批人) = "审批人"
        .TextMatrix(0, ColAdded.审批时间) = "审批时间"
        .TextMatrix(0, ColAdded.ID) = "ID"

        .Cell(flexcpAlignment, 0, ColAdded.类别, 0, .Cols - 1) = 4
        .ColWidth(ColAdded.选择) = 240
        .ColWidth(ColAdded.使用限量) = 650
        .ColWidth(ColAdded.类别) = 650
        .ColWidth(ColAdded.编码) = 1100
        .ColWidth(ColAdded.名称) = 2000
        .ColWidth(ColAdded.规格) = 1800
        .ColWidth(ColAdded.产地) = 2000
        .ColWidth(ColAdded.单位) = 500
        .ColWidth(ColAdded.说明) = 3000
        .ColWidth(ColAdded.审批人) = 900
        .ColWidth(ColAdded.审批时间) = 900
        .ColWidth(ColAdded.ID) = 0
        If mblnDel = False Then
            .ColHidden(ColAdded.选择) = True
        End If
        .ColHidden(ColAdded.ID) = True
    End With
    
    
    vsExist.ColHidden(ColAdded.ID) = True
    If mbln模板 Then
        vsExist.ColHidden(ColAdded.使用限量) = True
        vsExist.ColHidden(ColAdded.审批人) = True
        vsExist.ColHidden(ColAdded.审批时间) = True
    End If
    
    With vsList
        .Rows = 2
        .Cols = 9
        .TextMatrix(0, ColAdd.选择) = ""
        .TextMatrix(0, ColAdd.类别) = "类别"
        .TextMatrix(0, ColAdd.名称) = "名称"
        .TextMatrix(0, ColAdd.编码) = "编码"
        .TextMatrix(0, ColAdd.规格) = "规格"
        .TextMatrix(0, ColAdd.产地) = "产地"
        .TextMatrix(0, ColAdd.单位) = "单位"
        .TextMatrix(0, ColAdd.说明) = "说明"
        .TextMatrix(0, ColAdd.ID) = "ID"
        
        .Cell(flexcpAlignment, 0, ColAdded.类别, 0, ColAdded.说明) = 4
        .ColWidth(ColAdd.类别) = 650
        .ColWidth(ColAdd.编码) = 1100
        .ColWidth(ColAdd.名称) = 1700
        .ColWidth(ColAdd.规格) = 1300
        .ColWidth(ColAdd.产地) = 1500
        .ColWidth(ColAdd.单位) = 500
        .ColWidth(ColAdd.说明) = 1700
        .ColWidth(ColAdd.ID) = 0
        
        .ColHidden(ColAdd.ID) = True
        .Cell(flexcpChecked, 1, ColAdd.选择) = 1
    End With
End Sub

Private Function FillTree() As Boolean
'功能:装入收费类别和收费细目的所有分类到tvwMain_S
    '本程序中树节点比其它程序的KEY值多一个字符，即第二位的类别编码
    Dim i As Long
    Dim objNode As Node
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    
    Screen.MousePointer = vbHourglass
    LockWindowUpdate tvwMain_S.hWnd
    tvwMain_S.Nodes.Clear
    tvwMain_S.Sorted = False
    
    '显示分类
    strSQL = " Select 级, 类型, A.ID, 上级id, 名称" & _
             " From (Select Level As 级, 0 As 类型, ID, 上级id, '[' || 编码 || ']' || 名称 As 名称" & _
             "        From 收费分类目录 A" & _
             "        Start With 上级id Is Null" & _
             "        Connect By Prior ID = 上级id) A," & _
             "      (Select Distinct (ID) ID" & _
             "        From 收费分类目录 A" & _
             "        Start With ID In (Select Distinct A.分类id From 收费项目目录 A,保险支付项目 D " & _
             "                          Where A.ID = D.收费细目ID And D.要求审批 = 1 And (A.站点=[1] Or A.站点 is Null))" & _
             "        Connect By Prior 上级id = ID) B" & _
             " Where a.ID = B.ID" & _
             " Union"
              
    strSQL = strSQL & _
             " Select 0 As 级,类型, To_Number('99999999' || 类型) As ID, -null As 上级id," & _
             "        Chr(13) || Decode(类型, 1, '西成药', 2, '中成药', 3, '中草药', 7, '卫生材料') As 名称" & _
             " From 诊疗分类目录 " & _
             " Where Instr(',1,2,3,7,', ',' || 类型 || ',') > 0" & _
             " Union"
    
    strSQL = strSQL & _
             " Select 级, 类型, A.ID, 上级id, 名称" & _
             " From (Select Level As 级, 类型, ID As ID, Nvl(上级id, To_Number('99999999' || 类型)) As 上级id," & _
             "               '[' || 编码 || ']' || 名称 As 名称" & _
             "        From 诊疗分类目录" & _
             "        Where Instr(',1,2,3,7,', ',' || 类型 || ',') > 0" & _
             "        Start With 上级id Is Null" & _
             "        Connect By Prior ID = 上级id) A," & _
             "      (Select Distinct ID" & _
             "        From 诊疗分类目录" & _
             "        Start With ID In (Select Distinct (B.分类id) 分类id" & _
             "                          From 收费项目目录 A, 诊疗项目目录 B, 药品规格 C,保险支付项目 D" & _
             "                          Where A.ID = C.药品id And B.ID = C.药名id AND A.ID = D.收费细目ID And D.要求审批 = 1" & _
             "                                 And (A.站点=[1] Or A.站点 is Null))" & _
             "        Connect By Prior 上级id = ID) B" & _
             " Where a.ID = B.ID"

    On Error GoTo errHandle
    'by lesfeng 2010-03-06 性能绑定
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, gstrNodeNo)
    For i = 1 To rsTmp.RecordCount
        If IsNull(rsTmp!上级ID) Then
            Set objNode = tvwMain_S.Nodes.Add(, , "_" & rsTmp!ID, rsTmp!名称, "RootS", "Exp")
        Else
            Set objNode = tvwMain_S.Nodes.Add("_" & rsTmp!上级ID, 4, "_" & rsTmp!ID, rsTmp!名称, "RootS", "Exp")
        End If
        objNode.Tag = rsTmp!类型 '存放分类类型:0-非药品和卫材,1-西成药,2-中成药,3-中草药,7-卫生材料
        objNode.ExpandedImage = "Exp"
        rsTmp.MoveNext
    Next
    If tvwMain_S.Nodes.Count > 0 Then
        tvwMain_S.Nodes(1).Expanded = True
        If tvwMain_S.Nodes(1).Children > 0 Then
            tvwMain_S.Nodes(1).Child.Selected = True
        Else
            tvwMain_S.Nodes(1).Selected = True
        End If
        tvwMain_S.SelectedItem.EnsureVisible
        Call tvwMain_S_NodeClick(tvwMain_S.SelectedItem)
    End If
    On Error Resume Next
    If Not tvwMain_S.Nodes.Item("_999999991") Is Nothing Then
        If tvwMain_S.Nodes.Item("_999999991").Children = 0 Then tvwMain_S.Nodes.Remove "_999999991"
    End If
    If Not tvwMain_S.Nodes.Item("_999999992") Is Nothing Then
        If tvwMain_S.Nodes.Item("_999999992").Children = 0 Then tvwMain_S.Nodes.Remove "_999999992"
    End If
    If Not tvwMain_S.Nodes.Item("_999999993") Is Nothing Then
        If tvwMain_S.Nodes.Item("_999999993").Children = 0 Then tvwMain_S.Nodes.Remove "_999999993"
    End If
    If Not tvwMain_S.Nodes.Item("_999999997") Is Nothing Then
        If tvwMain_S.Nodes.Item("_999999997").Children = 0 Then tvwMain_S.Nodes.Remove "_999999997"
    End If
    FillTree = True
    Screen.MousePointer = 0
    LockWindowUpdate 0
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Screen.MousePointer = 0
End Function

Private Function FillList(str类型 As String, str分类 As String) As Boolean
    Dim strSQL  As String, str类别 As String
    Dim rsTemp As New ADODB.Recordset
    Dim lngRow As Long
    Dim int类别 As Integer
    Dim bln分类 As Boolean
    Dim lng编码 As Long
    
    If mbln模板 = True Then
        If lvw模板.SelectedItem Is Nothing Then
            lng编码 = 0
        ElseIf UCase(Mid(lvw模板.SelectedItem.Key, 2)) = "ADD" Then
            lng编码 = 0
        Else
            lng编码 = Mid(lvw模板.SelectedItem.Key, 2)
        End If
    End If
    
    Select Case str类型
        Case 1
            str类别 = 5
        Case 2
            str类别 = 6
        Case 3
            str类别 = 7
        Case 7
            str类别 = 4
    End Select
    
    bln分类 = InStr(str分类, "99999") = 0
    Screen.MousePointer = 11
    If mbln模板 = False Then
        If str类型 = 0 Then
             strSQL = " Select C.名称 类别名称, A.类别, A.编码, A.名称, A.规格, A.产地, A.计算单位, A.说明, A.ID" & _
                      " From 收费项目目录 A," & _
                      "      (Select A.ID" & _
                      "        From 收费项目目录 A,保险支付项目 D" & _
                      "        Where A.分类id In (Select ID From 收费分类目录 Start With ID = [1] Connect By Prior ID = 上级id) And" & _
                      "              A.服务对象 In (2, 3) And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null) And" & _
                      "              A.ID = D.收费细目ID And D.险类 = [4] And D.要求审批 = 1" & _
                      "              And (A.站点=[5] Or A.站点 is Null)" & _
                      "        Minus" & _
                      "        Select 项目ID From 病人审批项目 A Where A.病人id = [2] And 主页ID = [3]) B, 收费项目类别 C" & _
                      " Where A.ID = B.ID And A.类别 = C.编码" & _
                      " Order By A.类别, A.编码"
            'by lesfeng 2010-03-06 性能绑定
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str分类, mlng病人ID, mlng主页ID, mlng险类, gstrNodeNo)
        Else
           If bln分类 = True Then
                str分类 = CStr(CLng(str分类))
                strSQL = " Select C.名称 类别名称,A.类别, A.编码, A.名称, A.规格, A.产地, A.计算单位, A.说明,A.ID" & _
                         " From 收费项目目录 A, " & _
                         " (SELECT A.ID " & _
                         "   FROM 收费项目目录 A,收费项目类别 B, 药品规格 D, 诊疗项目目录 E,保险支付项目 F" & _
                         "   Where A.类别 = B.编码 And A.ID = D.药品id And D.药名id = E.ID And" & _
                         "      E.分类id In (Select ID From 诊疗分类目录 Start With ID = [1] Connect By Prior ID = 上级id) And" & _
                         "      A.服务对象 In (2, 3) And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null) And" & _
                         "      A.类别=[2] AND A.ID = F.收费细目ID And F.险类 = [5] And F.要求审批 = 1" & _
                         "      And (A.站点=[6] Or A.站点 is Null) And (E.站点=[6] Or E.站点 is Null)" & _
                         "   Minus" & _
                         "   Select 项目id From 病人审批项目 A Where A.病人id = [3] And 主页ID = [4]) B, 收费项目类别 C" & _
                         " Where  A.ID = B.ID And A.类别 = C.编码" & _
                         " Order By A.编码"
                'by lesfeng 2010-03-06 性能绑定
                Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CLng(str分类), str类别, mlng病人ID, mlng主页ID, mlng险类, gstrNodeNo)
            Else
                strSQL = " Select C.名称 类别名称, A.类别, A.编码, A.名称, A.规格, A.产地, A.计算单位, A.说明, A.ID" & _
                         " From 收费项目目录 A, " & _
                         "      (Select A.ID" & _
                         "       From 收费项目目录 A,保险支付项目 D" & _
                         "       Where A.服务对象 In (2, 3) And" & _
                         "            (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null) And" & _
                         "             A.类别 = [1] And A.ID = D.收费细目ID And D.险类 = [4] And D.要求审批 = 1" & _
                         "             And (A.站点=[5] Or A.站点 is Null)" & _
                         "       Minus" & _
                         "       Select 项目id From 病人审批项目 Where 病人id = [2] And 主页ID = [3]) B, 收费项目类别 C" & _
                         "  Where A.ID = B.ID And A.类别 = C.编码" & _
                         " Order By A.编码"
                'by lesfeng 2010-03-06 性能绑定
                Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str类别, mlng病人ID, mlng主页ID, mlng险类, gstrNodeNo)
            End If
        End If
    Else
        If str类型 = 0 Then
             strSQL = " Select C.名称 类别名称, A.类别, A.编码, A.名称, A.规格, A.产地, A.计算单位, A.说明, A.ID" & _
                      " From 收费项目目录 A," & _
                      "      (Select Distinct A.ID" & _
                      "        From 收费项目目录 A,保险支付项目 D" & _
                      "        Where A.分类id In (Select ID From 收费分类目录 Start With ID = [1] Connect By Prior ID = 上级id) And" & _
                      "              A.服务对象 In (2, 3) And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null) And" & _
                      "              A.ID = D.收费细目ID And D.要求审批 = 1" & _
                      "             And (A.站点=[3] Or A.站点 is Null)" & _
                      "        Minus" & _
                      "        Select 项目ID From 审批项目模板 A Where A.编码 = [2]) B, 收费项目类别 C" & _
                      " Where A.ID = B.ID And A.类别 = C.编码" & _
                      " Order By A.类别, A.编码"
             'by lesfeng 2010-03-06 性能绑定
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str分类, lng编码, gstrNodeNo)
        Else
           If bln分类 = True Then
                str分类 = CStr(CLng(str分类))
                strSQL = " Select C.名称 类别名称,A.类别, A.编码, A.名称, A.规格, A.产地, A.计算单位, A.说明,A.ID" & _
                         " From 收费项目目录 A, " & _
                         " (SELECT Distinct A.ID " & _
                         "   FROM 收费项目目录 A,收费项目类别 B, 药品规格 D, 诊疗项目目录 E,保险支付项目 F" & _
                         "   Where A.类别 = B.编码 And A.ID = D.药品id And D.药名id = E.ID And" & _
                         "      E.分类id In (Select ID From 诊疗分类目录 Start With ID = [1] Connect By Prior ID = 上级id) And" & _
                         "      A.服务对象 In (2, 3) And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null) And" & _
                         "      A.类别=[2] AND A.ID = F.收费细目ID And F.要求审批 = 1" & _
                         "      And (A.站点=[4] Or A.站点 is Null) And (E.站点=[4] Or E.站点 is Null)" & _
                         "   Minus" & _
                         "   Select 项目id From 审批项目模板 A Where A.编码 = [3]) B, 收费项目类别 C" & _
                         " Where  A.ID = B.ID And A.类别 = C.编码" & _
                         " Order By A.编码"
                'by lesfeng 2010-03-06 性能绑定
                Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CLng(str分类), str类别, lng编码, gstrNodeNo)
            Else
                strSQL = " Select C.名称 类别名称, A.类别, A.编码, A.名称, A.规格, A.产地, A.计算单位, A.说明, A.ID" & _
                         " From 收费项目目录 A, " & _
                         "      (Select Distinct A.ID" & _
                         "       From 收费项目目录 A,保险支付项目 D" & _
                         "       Where A.服务对象 In (2, 3) And" & _
                         "            (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null) And" & _
                         "             A.类别 = [1] And A.ID = D.收费细目ID And D.要求审批 = 1" & _
                         "           And (A.站点=[3] Or A.站点 is Null)" & _
                         "       Minus" & _
                         "       Select 项目id From 审批项目模板 Where 编码 = [2]) B, 收费项目类别 C" & _
                         "  Where A.ID = B.ID And A.类别 = C.编码" & _
                         " Order By A.编码"
                'by lesfeng 2010-03-06 性能绑定
                Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str类别, lng编码, gstrNodeNo)
            End If
        End If
    End If
On Error GoTo errHandle
    
    vsList.Rows = 1
    lngRow = 1
    While Not rsTemp.EOF
        With vsList
            .Rows = lngRow + 1
            .TextMatrix(lngRow, ColAdd.编码) = rsTemp!编码
            .TextMatrix(lngRow, ColAdd.类别) = rsTemp!类别名称
            If str类别 <> rsTemp!类别名称 Then
                str类别 = rsTemp!类别名称
                int类别 = int类别 + 1
            End If
            .TextMatrix(lngRow, ColAdd.名称) = rsTemp!名称
            .TextMatrix(lngRow, ColAdd.规格) = rsTemp!规格 & ""
            .TextMatrix(lngRow, ColAdd.单位) = rsTemp!计算单位 & ""
            .TextMatrix(lngRow, ColAdd.产地) = rsTemp!产地 & ""
            .TextMatrix(lngRow, ColAdd.说明) = rsTemp!说明 & ""
            .TextMatrix(lngRow, ColAdd.ID) = rsTemp!ID
        End With
        lngRow = lngRow + 1
        rsTemp.MoveNext
    Wend
    If int类别 = 0 Or int类别 = 1 Then
        vsList.ColHidden(ColAdd.类别) = True
    Else
        vsList.ColHidden(ColAdd.类别) = False
    End If
    
    If str类型 = 0 Then
        vsList.ColHidden(ColAdd.产地) = True
    Else
        vsList.ColHidden(ColAdd.产地) = False
    End If
    
    If vsList.Rows > 1 Then
        vsList.Cell(flexcpChecked, 1, ColAdd.选择, vsList.Rows - 1, ColAdd.选择) = 2
    End If
    Screen.MousePointer = 0
    vsList.Editable = flexEDKbdMouse
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Screen.MousePointer = 0
End Function

Private Sub tabClass_Click()
    If tabClass.SelectedItem.Index <> 1 Then
        mrsExistItem.Filter = "类别='" & tabClass.SelectedItem.Tag & "'"
    Else
        mrsExistItem.Filter = 0
    End If
    Set vsExist.DataSource = mrsExistItem
    If tabClass.SelectedItem.Index <> 1 Then
        vsExist.ColHidden(ColAdded.类别) = True
    Else
        vsExist.ColHidden(ColAdded.类别) = False
    End If
    vsExist.ColHidden(ColAdded.ID) = True
    
    If InStr("中草药,中成药,西成药,材料", tabClass.SelectedItem.Tag) = 0 Then
        vsExist.ColHidden(ColAdded.产地) = True
    Else
         vsExist.ColHidden(ColAdded.产地) = False
    End If
    
    If mbln模板 Then
        vsExist.ColHidden(ColAdded.使用限量) = True
        vsExist.ColHidden(ColAdded.审批人) = True
        vsExist.ColHidden(ColAdded.审批时间) = True
    End If
    
    vsExist.TextMatrix(0, 0) = ""
    vsExist.ColWidth(ColAdded.选择) = 240
    If vsExist.Rows > 1 Then
        vsExist.Cell(flexcpChecked, 1, ColAdded.选择, vsExist.Rows - 1, ColAdded.选择) = 2
    End If
    
    vsExist.ColAlignment(ColAdded.编码) = flexAlignLeftCenter
    vsExist.Tag = tabClass.SelectedItem.Index
End Sub

Private Sub tvwMain_S_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And Not mbln模板 Then
        If lvw模板.Visible Then lvw模板.Visible = False
    End If
End Sub

Private Sub tvwMain_S_NodeClick(ByVal Node As MSComctlLib.Node)
    If vsList.Tag <> Node.Key Then
        Call FillList(Node.Tag, Mid(Node.Key, 2))
        vsList.Tag = Node.Key
    End If
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Trim(txtFind.Text) = "" Then Exit Sub
        If UCase(txtFind.Text) <> UCase(txtFind.Tag) Then
            Call FindItem(Trim(txtFind.Text))
            txtFind.Tag = txtFind.Text
            vsList.Tag = ""
        End If
    End If
End Sub

Private Sub vsExist_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = ColAdded.使用限量 Then
        On Error GoTo errHandle
        With vsExist
            gstrSQL = "Zl_病人审批项目_Update(" & mlng病人ID & "," & mlng主页ID & "," & .TextMatrix(Row, ColAdded.ID) & "," & Val(.TextMatrix(Row, Col)) & ")"
        End With
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Screen.MousePointer = 0
End Sub

Private Sub vsExist_EnterCell()
    If vsExist.Col = ColAdded.选择 Then
        vsExist.Editable = flexEDKbdMouse
    ElseIf vsExist.Col = ColAdded.使用限量 Then
        vsExist.Editable = flexEDKbdMouse
    Else
        vsExist.Editable = flexEDNone
    End If
End Sub

Private Sub vsExist_KeyDown(KeyCode As Integer, Shift As Integer)
    If vsExist.Row > 0 Then
        If KeyCode = vbKeyDelete Then
            KeyCode = 0
            If vsExist.Cell(flexcpChecked, vsExist.Row, ColAdded.选择) = 1 Then
                If mblnDelPrv = True Then
                    Call cmdDelete_Click
                End If
            Else
                vsExist.TextMatrix(vsExist.Row, ColAdded.使用限量) = ""
                Call vsExist_AfterEdit(vsExist.Row, ColAdded.使用限量)
            End If
        End If
    End If
End Sub

Private Sub vsExist_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Not KeyAscii = vbKeyReturn And Col = ColAdded.使用限量 Then
        If InStr("0123456789." & Chr(8) & Chr(27), Chr(KeyAscii)) = 0 Then
            KeyAscii = 0: Exit Sub
        End If
    End If
End Sub

Private Sub vsExist_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        If vsExist.Row < vsExist.Rows - 1 Then
            vsExist.Row = vsExist.Row + 1
            vsExist.Col = ColAdded.使用限量
            vsExist.ShowCell vsExist.Row, vsExist.Col
        Else
            cmdDelete.SetFocus
        End If
    ElseIf KeyAscii = vbKeySpace Then
        KeyAscii = 0
        vsExist.Editable = flexEDKbdMouse
        vsExist.Cell(flexcpChecked, vsExist.Row, ColAdded.选择) = IIf(vsExist.Cell(flexcpChecked, vsExist.Row, ColAdded.选择) = 1, 2, 1)
        vsExist.Editable = flexEDNone
    End If
End Sub

Private Sub vsExist_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And Not mbln模板 Then
        If lvw模板.Visible Then lvw模板.Visible = False
    End If
End Sub

Private Sub vsList_EnterCell()
    If vsList.Col = ColAdd.选择 Then
        vsList.Editable = flexEDKbdMouse
    Else
        vsList.Editable = flexEDNone
    End If
End Sub

Private Sub vsList_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        If vsList.Row < vsList.Rows - 1 Then
            vsList.Row = vsList.Row + 1
            Do Until vsList.RowHidden(vsList.Row) = False
                vsList.Row = vsList.Row + 1
            Loop
            vsList.Col = ColAdd.选择
            vsList.ShowCell vsList.Row, vsList.Col
        End If
    ElseIf KeyAscii = vbKeySpace Then
        KeyAscii = 0
        vsList.Editable = flexEDKbdMouse
        vsList.Cell(flexcpChecked, vsList.Row, ColAdd.选择) = IIf(vsList.Cell(flexcpChecked, vsList.Row, ColAdd.选择) = 1, 2, 1)
        vsList.Editable = flexEDNone
    End If
End Sub

Private Function ReadExistsItem(lng病人ID As Long, lng主页ID As Long) As Boolean
    Dim strSQL As String
    Dim lngRow As Long
    Dim strClass As String, strOld As String
    Dim arrClass As Variant
    Dim blnClass As Boolean
    Dim objTab As MSComctlLib.Tab
    Dim i As Integer
    
    strSQL = " Select NULL,B.使用限量,C.名称 类别, A.编码, A.名称, A.规格, A.产地, A.计算单位, A.说明,B.审批人,TRUNC(B.审批时间) 审批时间,A.ID" & _
             " From 收费项目目录 A,病人审批项目 B, 收费项目类别 C " & _
             " Where A.类别 = C.编码 And A.ID=B.项目ID AND B.病人ID=[1] AND B.主页ID=[2]" & _
             "       And (A.站点=[3] Or A.站点 is Null)" & _
             " order by 类别,编码"
On Error GoTo errHandle
    'by lesfeng 2010-03-06 性能绑定
    Set mrsExistItem = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng病人ID, lng主页ID, gstrNodeNo)
    mlngCount = mrsExistItem.RecordCount
    Set vsExist.DataSource = mrsExistItem
    
    If mrsExistItem.RecordCount = 0 Then
        If mblnDel = True Then
            ReadExistsItem = False
            Exit Function
        Else
            With vsExist
                .Cell(flexcpAlignment, 0, ColAdded.类别, 0, .Cols - 1) = 4
                .ColWidth(ColAdded.选择) = 240
                .ColWidth(ColAdded.类别) = 650
                .ColWidth(ColAdded.编码) = 1100
                .ColWidth(ColAdded.名称) = 1700
                .ColWidth(ColAdded.规格) = 1300
                .ColWidth(ColAdded.产地) = 1500
                .ColWidth(ColAdded.单位) = 500
                .ColWidth(ColAdded.说明) = 1700
                .ColWidth(ColAdded.ID) = 0
            End With
        End If
    End If
    
    While Not mrsExistItem.EOF
        If mrsExistItem!类别 <> strOld Then
            strClass = strClass & "," & mrsExistItem!类别
            strOld = mrsExistItem!类别
        End If
        mrsExistItem.MoveNext
    Wend
    
    For i = tabClass.Tabs.Count To 2 Step -1
        tabClass.Tabs.Remove i
    Next
    
    arrClass = Split(Mid(strClass, 2), ",")
    
    If UBound(arrClass) > 0 Then
        tabClass.Visible = True
        tabClass.ZOrder
        Call Form_Resize
        For i = 0 To UBound(arrClass)
            If i < 9 Then
                '用Alt快捷键焦点无法处理
                Set objTab = tabClass.Tabs.Add(, arrClass(i), arrClass(i) & "(&" & i + 1 & ")")
            Else
                Set objTab = tabClass.Tabs.Add(, arrClass(i), arrClass(i))
            End If
            objTab.Tag = arrClass(i)
        Next
    Else
        tabClass.Visible = False
    End If
    vsExist.ColWidth(ColAdded.ID) = 0
    vsExist.ColHidden(ColAdded.ID) = True
    
    vsExist.TextMatrix(0, 0) = ""
    vsExist.ColWidth(ColAdded.选择) = 240
    If vsExist.Rows > 1 Then
        vsExist.Cell(flexcpChecked, 1, ColAdded.选择, vsExist.Rows - 1, ColAdded.选择) = 2
    End If

    vsExist.ColAlignment(ColAdded.编码) = flexAlignLeftCenter
    If vsExist.Tag <> "" Then
        If vsExist.Tag < tabClass.Tabs.Count Then
            Set tabClass.SelectedItem = tabClass.Tabs.Item(Int(vsExist.Tag))
        End If
        Call tabClass_Click
    End If
    
    If vsExist.Col = 1 Then
        vsExist.Editable = flexEDKbdMouse
    End If
    Call View状态
    Call Form_Resize
    vsExist.Editable = flexEDKbdMouse
    ReadExistsItem = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    ReadExistsItem = False
    Call SaveErrLog
End Function

Private Sub ReadTempletItem(lng编码 As Long)
    Dim strSQL As String
    Dim lngRow As Long
    Dim strClass As String, strOld As String
    Dim arrClass As Variant
    Dim blnClass As Boolean
    Dim objTab As MSComctlLib.Tab
    Dim i As Integer
             
             
    strSQL = "Select NULL,NULL,C.名称 类别, A.编码, A.名称, A.规格, A.产地, A.计算单位, A.说明,Null 审批人,Null 审批时间,A.Id " & _
             "From 收费项目目录 A,审批项目模板 B, 收费项目类别 C " & _
             "Where A.类别 = C.编码 And A.ID = B.项目ID And B.编码 = [1] " & _
             "      And (A.站点=[2] Or A.站点 is Null)" & _
             "Order By 类别,编码"
             
On Error GoTo errHandle
    'by lesfeng 2010-03-06 性能绑定
    Set mrsExistItem = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng编码, gstrNodeNo)
    mlngCount = mrsExistItem.RecordCount
    Set vsExist.DataSource = mrsExistItem
    
    If mrsExistItem.RecordCount = 0 Then
        If mblnDel = True Then
            Unload Me
            Exit Sub
        Else
            With vsExist
                .Cell(flexcpAlignment, 0, ColAdded.类别, 0, .Cols - 1) = 4
                .ColWidth(ColAdded.选择) = 240
                .ColWidth(ColAdded.类别) = 650
                .ColWidth(ColAdded.编码) = 1100
                .ColWidth(ColAdded.名称) = 1700
                .ColWidth(ColAdded.规格) = 1300
                .ColWidth(ColAdded.产地) = 1500
                .ColWidth(ColAdded.单位) = 500
                .ColWidth(ColAdded.说明) = 1700
            End With
        End If
    End If
    
    While Not mrsExistItem.EOF
        If mrsExistItem!类别 <> strOld Then
            strClass = strClass & "," & mrsExistItem!类别
            strOld = mrsExistItem!类别
        End If
        mrsExistItem.MoveNext
    Wend
    
    For i = tabClass.Tabs.Count To 2 Step -1
        tabClass.Tabs.Remove i
    Next
    
    arrClass = Split(Mid(strClass, 2), ",")
    
    If UBound(arrClass) > 0 Then
        tabClass.Visible = True
        tabClass.ZOrder
        Call Form_Resize
        For i = 0 To UBound(arrClass)
            If i < 9 Then
                '用Alt快捷键焦点无法处理
                Set objTab = tabClass.Tabs.Add(, arrClass(i), arrClass(i) & "(&" & i + 1 & ")")
            Else
                Set objTab = tabClass.Tabs.Add(, arrClass(i), arrClass(i))
            End If
            objTab.Tag = arrClass(i)
        Next
    Else
        tabClass.Visible = False
    End If
    
    vsExist.ColHidden(ColAdded.ID) = True
    vsExist.ColHidden(ColAdded.使用限量) = True
    vsExist.ColHidden(ColAdded.审批人) = True
    vsExist.ColHidden(ColAdded.审批时间) = True

    vsExist.TextMatrix(0, 0) = ""
    vsExist.ColWidth(ColAdded.选择) = 240
    If vsExist.Rows > 1 Then
        vsExist.Cell(flexcpChecked, 1, ColAdded.选择, vsExist.Rows - 1, ColAdded.选择) = 2
    End If
    
    vsExist.ColAlignment(ColAdded.编码) = flexAlignLeftCenter
    If vsExist.Tag <> "" Then
        If vsExist.Tag < tabClass.Tabs.Count Then
            Set tabClass.SelectedItem = tabClass.Tabs.Item(Int(vsExist.Tag))
        End If
        Call tabClass_Click
    End If
    
    If vsExist.Col = 1 Then
        vsExist.Editable = flexEDKbdMouse
    End If
    Call View状态
    Call Form_Resize
    vsExist.Editable = flexEDKbdMouse
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
'    Resume
End Sub

Private Sub FindItem(strWhere As String)
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim str类别 As String
    Dim int类别 As Integer, lngRow As Long, lng编码 As Long
    
    If mbln模板 = True Then
        If UCase(Mid(lvw模板.SelectedItem.Key, 2)) = "ADD" Then Exit Sub
        lng编码 = Mid(lvw模板.SelectedItem.Key, 2)
    End If
    
On Error GoTo errHandle
    Screen.MousePointer = 11
    If mbln模板 = False Then
        If IsNumeric(Trim(strWhere)) Then
            
            strWhere = gstrLike & strWhere & "%"
        
            strSQL = " Select C.名称 类别名称, A.类别, A.编码, A.名称, A.规格, A.产地, A.计算单位, A.说明, A.ID" & _
                     " From 收费项目目录 A," & _
                     "      (Select A.ID" & _
                     "       From 收费项目目录 A,保险支付项目 D" & _
                     "       Where (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null) AND A.服务对象 IN (2,3) And  " & _
                     "             A.ID = D.收费细目ID And D.险类 = [4] And D.要求审批 = 1 And A.编码 Like [1]" & _
                     "             And (A.站点=[5] Or A.站点 is Null)" & _
                     "       Minus" & _
                     "       Select A.项目id From 病人审批项目 A Where 病人id =[2] AND 主页ID = [3]) B, 收费项目类别 C" & _
                     " Where A.ID = B.ID And A.类别 = C.编码" & _
                     " ORDER BY 类别,编码"
            'by lesfeng 2010-03-06 性能绑定
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strWhere, mlng病人ID, mlng主页ID, mlng险类, gstrNodeNo)
            
        ElseIf zlCommFun.IsCharAlpha(Trim(txtFind.Text)) Then
        '只是字母时查询简码
            strWhere = gstrLike & strWhere & "%"
            
            strSQL = " Select C.名称 类别名称, A.类别, A.编码, A.名称, A.规格, A.产地, A.计算单位, A.说明, A.ID" & _
                     " From 收费项目目录 A," & _
                     "      (Select 收费细目id" & _
                     "        From 收费项目别名" & _
                     "        Where " & IIf(gbytCode + 1 = 3, "", "码类 = [1] And") & " 简码 Like [2]" & _
                     "        Group By 收费细目id" & _
                     "        Minus" & _
                     "        Select A.项目id From 病人审批项目 A Where 病人id = [3] And 主页ID = [4]) B, 收费项目类别 C,保险支付项目 D" & _
                     " Where A.ID = B.收费细目id And A.类别 = C.编码 And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null) And" & _
                     "       A.服务对象 In (2, 3) And A.ID = D.收费细目ID And D.险类 = [5] And D.要求审批 = 1" & _
                     "       And (A.站点=[6] Or A.站点 is Null)" & _
                     " ORDER BY 类别,编码"
            'by lesfeng 2010-03-06 性能绑定
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, gbytCode + 1, UCase(strWhere), mlng病人ID, mlng主页ID, mlng险类, gstrNodeNo)
        Else
            strWhere = gstrLike & strWhere & "%"
            
            strSQL = " Select C.名称 类别名称, A.类别, A.编码, A.名称, A.规格, A.产地, A.计算单位, A.说明, A.ID" & _
                     " From 收费项目目录 A," & _
                     "      (Select 收费细目id" & _
                     "        From 收费项目别名" & _
                     "        Where 码类=1 AND 名称 Like [1]" & _
                     "        Group By 收费细目id" & _
                     "        Minus" & _
                     "        Select A.项目id From 病人审批项目 A Where 病人id = [2] And 主页ID = [3]) B, 收费项目类别 C,保险支付项目 D" & _
                     " Where A.ID = B.收费细目id And A.类别 = C.编码 And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null) And" & _
                     "       A.服务对象 In (2, 3) And A.ID = D.收费细目ID And D.险类 = [4] And D.要求审批 = 1" & _
                     "       And (A.站点=[5] Or A.站点 is Null)" & _
                     " ORDER BY 类别,编码"
            'by lesfeng 2010-03-06 性能绑定
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UCase(strWhere), mlng病人ID, mlng主页ID, mlng险类, gstrNodeNo)
        End If
    Else
         If IsNumeric(Trim(strWhere)) Then
            
            strWhere = gstrLike & strWhere & "%"
        
            strSQL = " Select C.名称 类别名称, A.类别, A.编码, A.名称, A.规格, A.产地, A.计算单位, A.说明, A.ID" & _
                     " From 收费项目目录 A," & _
                     "      (Select Distinct A.ID" & _
                     "       From 收费项目目录 A,保险支付项目 D" & _
                     "       Where (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null) AND A.服务对象 IN (2,3) And  " & _
                     "             A.ID = D.收费细目ID And D.要求审批 = 1 And A.编码 Like [1]" & _
                     "             And (A.站点=[3] Or A.站点 is Null)" & _
                     "       Minus" & _
                     "       Select A.项目id From 审批项目模板 A Where 编码 =[2] ) B, 收费项目类别 C" & _
                     " Where A.ID = B.ID And A.类别 = C.编码" & _
                     " ORDER BY 类别,编码"
            'by lesfeng 2010-03-06 性能绑定
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strWhere, lng编码, gstrNodeNo)
            
        ElseIf zlCommFun.IsCharAlpha(Trim(txtFind.Text)) Then
        '只是字母时查询简码
            strWhere = gstrLike & strWhere & "%"
            
            strSQL = " Select Distinct C.名称 类别名称, A.类别, A.编码, A.名称, A.规格, A.产地, A.计算单位, A.说明, A.ID" & _
                     " From 收费项目目录 A," & _
                     "      (Select 收费细目id" & _
                     "        From 收费项目别名" & _
                     "        Where " & IIf(gbytCode + 1 = 3, "", "码类 = [1] And") & " 简码 Like [2]" & _
                     "        Group By 收费细目id" & _
                     "        Minus" & _
                     "        Select A.项目id From 审批项目模板 A Where 编码 = [3]) B, 收费项目类别 C,保险支付项目 D" & _
                     " Where A.ID = B.收费细目id And A.类别 = C.编码 And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null) And" & _
                     "       A.服务对象 In (2, 3) And A.ID = D.收费细目ID And D.要求审批 = 1" & _
                     "       And (A.站点=[4] Or A.站点 is Null)" & _
                     " ORDER BY 类别,编码"
            'by lesfeng 2010-03-06 性能绑定
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, gbytCode + 1, UCase(strWhere), lng编码, gstrNodeNo)
        Else
            strWhere = gstrLike & strWhere & "%"
            
            strSQL = " Select Distinct C.名称 类别名称, A.类别, A.编码, A.名称, A.规格, A.产地, A.计算单位, A.说明, A.ID" & _
                     " From 收费项目目录 A," & _
                     "      (Select 收费细目id" & _
                     "        From 收费项目别名" & _
                     "        Where 码类=1 AND 名称 Like [1]" & _
                     "        Group By 收费细目id" & _
                     "        Minus" & _
                     "        Select A.项目id From 审批项目模板 A Where 编码= [2] ) B, 收费项目类别 C,保险支付项目 D" & _
                     " Where A.ID = B.收费细目id And A.类别 = C.编码 And (A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.撤档时间 Is Null) And" & _
                     "       A.服务对象 In (2, 3) And A.ID = D.收费细目ID And D.要求审批 = 1" & _
                     "       And (A.站点=[3] Or A.站点 is Null)" & _
                     " ORDER BY 类别,编码"
            'by lesfeng 2010-03-06 性能绑定
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UCase(strWhere), lng编码, gstrNodeNo)
        End If
    End If
    vsList.Rows = 1
    lngRow = 1
    While Not rsTmp.EOF
        With vsList
            .Rows = lngRow + 1
            .TextMatrix(lngRow, ColAdd.编码) = rsTmp!编码
            .TextMatrix(lngRow, ColAdd.类别) = rsTmp!类别名称
            If str类别 <> rsTmp!类别名称 Then
                str类别 = rsTmp!类别名称
                int类别 = int类别 + 1
            End If
            .TextMatrix(lngRow, ColAdd.名称) = rsTmp!名称
            .TextMatrix(lngRow, ColAdd.规格) = rsTmp!规格 & ""
            .TextMatrix(lngRow, ColAdd.单位) = rsTmp!计算单位 & ""
            .TextMatrix(lngRow, ColAdd.产地) = rsTmp!产地 & ""
            .TextMatrix(lngRow, ColAdd.说明) = rsTmp!说明 & ""
            .TextMatrix(lngRow, ColAdd.ID) = rsTmp!ID
        End With
        lngRow = lngRow + 1
        rsTmp.MoveNext
    Wend
    If int类别 = 0 Or int类别 = 1 Then
        vsList.ColHidden(ColAdd.类别) = True
    Else
        vsList.ColHidden(ColAdd.类别) = False
    End If
    If vsList.Rows > 1 Then
        vsList.Cell(flexcpChecked, 1, ColAdd.选择, vsList.Rows - 1, ColAdd.选择) = 2
    End If
    Screen.MousePointer = 0
    Exit Sub
    
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
     Screen.MousePointer = 0
End Sub

Public Sub ExamineEdit(lng病人ID As Long, lng主页ID As Long, lng险类 As Long, Optional blnDel As Boolean = False, Optional bln模板 As Boolean = False)
    
    RestoreWinState Me, App.ProductName
    
    mlng病人ID = lng病人ID
    mlng主页ID = lng主页ID
    mlng险类 = lng险类
    mblnDel = blnDel
    mblnDelPrv = InStr(frmManageExamine.mstrPrivs, "删除审批项目")
    mbln模板 = bln模板
    
    Call InitForm
    
    If mblnDel = False Then
        Call FillTree
        cmdDelete.Caption = "删除(&D)"
        cmdAdd.Caption = "增加(&A)"
        Me.Caption = "病人费用审批项目编辑(当前用户：" & UserInfo.姓名 & ")"
    Else
        cmdDelete.Caption = "确定(&O)"
        cmdAdd.Caption = "取消(&C)"
        Me.Caption = "病人费用审批项目编辑(当前用户：" & UserInfo.姓名 & ")"
        
    End If
    If mbln模板 = False Then
        lvw模板.Visible = False
        cmd调用.Visible = mblnDel = False
        lbl模板.Visible = False
        lbl项目.Visible = False
        lbl分类.Visible = False
        Me.BackColor = &H8000000F
        cmdSave.Visible = InStr(frmManageExamine.mstrPrivs, "模板管理") And mblnDel = False
        If mblnDelPrv = False Then
            cmdDelete.Visible = False
            cmdAdd.Left = cmdDelete.Left
            ImgFind.Left = cmdAdd.Left + cmdAdd.Width + 50
            txtFind.Left = ImgFind.Left + ImgFind.Width + 50
        End If
        
        If ReadExistsItem(mlng病人ID, mlng主页ID) = False Then Exit Sub
        Call ReadExistsTemplet
        If mblnDel = True And vsExist.Rows < 2 Then
            MsgBox "该病人没有设置费用审批项目!", vbInformation, gstrSysName
            Exit Sub
        End If
    Else
        Me.BackColor = &H808080
        Me.Caption = "审批项目模板编辑(当前用户：" & UserInfo.姓名 & ")"
        lbl模板.Visible = True
        lbl项目.Visible = True
        lbl分类.Visible = True
        Call ReadExistsTemplet
    End If
    Call Form_Resize
    frmExamineEdit.Show 1, frmManageExamine
End Sub

Private Sub vsList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 And Not mbln模板 Then
        If lvw模板.Visible Then lvw模板.Visible = False
    End If
End Sub

