VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Begin VB.Form frmPIVASortSet 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "排序设置"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5565
   Icon            =   "frmPIVASortSet.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   5565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdAddLeft 
      Height          =   375
      Index           =   1
      Left            =   2400
      Picture         =   "frmPIVASortSet.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton cmdDownUp 
      Height          =   375
      Index           =   1
      Left            =   5040
      Picture         =   "frmPIVASortSet.frx":0596
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3000
      TabIndex        =   7
      Top             =   4080
      Width           =   1100
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4200
      TabIndex        =   6
      Top             =   4080
      Width           =   1100
   End
   Begin VB.Frame Frame1 
      Height          =   75
      Left            =   0
      TabIndex        =   5
      Top             =   3840
      Width           =   6345
   End
   Begin VB.CommandButton cmdAddLeft 
      Height          =   375
      Index           =   0
      Left            =   2400
      Picture         =   "frmPIVASortSet.frx":0B20
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1200
      Width           =   375
   End
   Begin VB.CommandButton cmdDownUp 
      Height          =   375
      Index           =   0
      Left            =   5040
      Picture         =   "frmPIVASortSet.frx":10AA
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1200
      Width           =   375
   End
   Begin VB.Frame fraBottom 
      Height          =   75
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   6345
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfColAll 
      Height          =   2970
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   2100
      _cx             =   3704
      _cy             =   5239
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   14737632
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483647
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   0
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   0
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   8
      Cols            =   1
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmPIVASortSet.frx":1634
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
      Ellipsis        =   1
      ExplorerBar     =   0
      PicturesOver    =   -1  'True
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
      WallPaperAlignment=   1
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfColSelect 
      Height          =   2970
      Left            =   2880
      TabIndex        =   2
      Top             =   840
      Width           =   2100
      _cx             =   3704
      _cy             =   5239
      Appearance      =   0
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   14737632
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483647
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   0
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   0
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   1
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmPIVASortSet.frx":16A6
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
      Ellipsis        =   1
      ExplorerBar     =   0
      PicturesOver    =   -1  'True
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
      WallPaperAlignment=   1
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   120
      Picture         =   "frmPIVASortSet.frx":16FC
      Top             =   120
      Width           =   360
   End
   Begin VB.Label Label1 
      Caption         =   "设置输液单在界面的显示顺序，请从左边列表中选择，并在右边列表中设置排列顺序"
      Height          =   420
      Left            =   600
      TabIndex        =   8
      Top             =   120
      Width           =   4860
   End
End
Attribute VB_Name = "frmPIVASortSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub SetColSort(ByVal intType As Integer, ByVal intRow As Integer)
    '设置列上下顺序
    'intType：0-向下；1-向上
    'intRow：行号
    Dim strItem, strNextItem As String

    With vsfColSelect
        '不处理的情况
        If intRow = 0 Then Exit Sub
        If intType = 0 And intRow = .rows - 1 Then Exit Sub
        If intType = 1 And intRow = 1 Then Exit Sub
        
        .Redraw = flexRDNone
        
        '先记录当前行和要移动的行的内容
        strItem = .TextMatrix(intRow, 0)
        strNextItem = .TextMatrix(IIf(intType = 0, intRow + 1, intRow - 1), 0)
        
        '交换当前行和替换行的内容
        .TextMatrix(intRow, 0) = strNextItem
        .TextMatrix(IIf(intType = 0, intRow + 1, intRow - 1), 0) = strItem
        
        '当前行要跟着变化
        .Row = IIf(intType = 0, intRow + 1, intRow - 1)
        
        .Redraw = flexRDDirect
        
    End With
    
End Sub

Private Sub SetColAddCacel(ByVal vsfOut As VSFlexGrid, ByVal vsfIn As VSFlexGrid, ByVal intRow As Integer)
    '设置列加入/取消排序队列
    'intType：0-加入；1-取消
    'intRow：行号
    Dim strItem As String
    
    '先记录原表格的项目内容，再删除
    With vsfOut
        If intRow = 0 Then Exit Sub
        
        strItem = .TextMatrix(intRow, 0)
        .RemoveItem intRow
    End With
        
    '加入到新表格最后一行
    With vsfIn
        .rows = .rows + 1
        .TextMatrix(.rows - 1, 0) = strItem
        .Row = .rows - 1
    End With
End Sub

Private Sub cmdAddLeft_Click(Index As Integer)
    If Index = 0 Then
        '加入排序
        Call SetColAddCacel(vsfColAll, vsfColSelect, vsfColAll.Row)
    Else
        '取消排序
        Call SetColAddCacel(vsfColSelect, vsfColAll, vsfColSelect.Row)
    End If
End Sub

Private Sub cmdCancel_Click()
     Unload Me
End Sub

Private Sub cmdDownUp_Click(Index As Integer)
    Call SetColSort(Index, vsfColSelect.Row)
End Sub


Private Sub cmdOk_Click()
    Dim strItem1, strItem2 As String
    Dim i As Integer

    With vsfColAll
        For i = 1 To .rows - 1
            strItem1 = IIf(strItem1 = "", "", strItem1 & ",") & .TextMatrix(i, 0)
        Next
    End With
    
    With vsfColSelect
        For i = 1 To .rows - 1
            strItem2 = IIf(strItem2 = "", "", strItem2 & ",") & .TextMatrix(i, 0)
        Next
    End With
    
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\输液配置中心管理", "输液单排序", strItem1 & "|" & strItem2)
    
    Unload Me
End Sub


Private Sub Form_Load()
    Dim strSort As String
    Dim strDefault As String
    Dim i As Integer
    Dim strTmp As String
    
    strDefault = "病区,科室,床号,配药批次,姓名,瓶签号,执行时间"
    strSort = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "输液配置中心管理", "输液单排序", "")
    
    If strSort = "" Then strSort = strDefault
    
    '排序字段合法性检查，只要存在不匹配的则取默认字段列表
    strTmp = Replace(strSort, "|", ",")
    For i = 0 To UBound(Split(strTmp, ","))
        If Split(strTmp, ",")(i) <> "" Then
            If InStr(1, "," & strDefault & ",", "," & Split(strTmp, ",")(i) & ",") = 0 Then
                strSort = strDefault
                Exit For
            End If
        End If
    Next
    
    If InStr(1, strSort, "|") = 0 Then strSort = strSort & "|"
    
    With vsfColAll
        .rows = 1
        For i = 0 To UBound(Split(Split(strSort, "|")(0), ","))
            .rows = .rows + 1
            .TextMatrix(.rows - 1, 0) = Split(Split(strSort, "|")(0), ",")(i)
        Next
    End With
    
    With vsfColSelect
        .rows = 1
        For i = 0 To UBound(Split(Split(strSort, "|")(1), ","))
            .rows = .rows + 1
            .TextMatrix(.rows - 1, 0) = Split(Split(strSort, "|")(1), ",")(i)
        Next
    End With
End Sub

