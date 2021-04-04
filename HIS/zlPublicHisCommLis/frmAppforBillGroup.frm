VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAppforBillGroup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "申请单下属分组设置"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12255
   Icon            =   "frmAppforBillGroup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   12255
   StartUpPosition =   1  '所有者中心
   Begin VB.ComboBox cboGroupSel 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5940
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   150
      Width           =   2025
   End
   Begin VB.PictureBox picRight 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4365
      Left            =   4410
      ScaleHeight     =   4365
      ScaleWidth      =   4995
      TabIndex        =   2
      Top             =   750
      Width           =   4995
      Begin VB.PictureBox picLast 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   60
         MouseIcon       =   "frmAppforBillGroup.frx":6852
         MousePointer    =   99  'Custom
         Picture         =   "frmAppforBillGroup.frx":69A4
         ScaleHeight     =   360
         ScaleWidth      =   360
         TabIndex        =   6
         ToolTipText     =   "上一页"
         Top             =   1440
         Width           =   360
      End
      Begin VB.PictureBox picFindNext 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   90
         MouseIcon       =   "frmAppforBillGroup.frx":708E
         MousePointer    =   99  'Custom
         Picture         =   "frmAppforBillGroup.frx":71E0
         ScaleHeight     =   360
         ScaleWidth      =   360
         TabIndex        =   5
         ToolTipText     =   "下一页"
         Top             =   1890
         Width           =   360
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFList 
         Height          =   1995
         Left            =   720
         TabIndex        =   3
         Top             =   660
         Width           =   2895
         _cx             =   5106
         _cy             =   3519
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
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
         BackColorSel    =   16772055
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483635
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   2
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   3
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   350
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   -1  'True
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
         ShowComboButton =   0
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
         Begin VB.Label lblShow 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "1111111"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   660
            TabIndex        =   11
            Top             =   1470
            Visible         =   0   'False
            Width           =   840
         End
      End
      Begin VB.Label lblShortCaptionItem 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "已分组项目(点击""调整顺序""之后可拖动改变顺序)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   8
         Top             =   30
         Width           =   5280
      End
   End
   Begin VB.PictureBox picLeft 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   0
      ScaleHeight     =   4455
      ScaleWidth      =   3225
      TabIndex        =   0
      Top             =   630
      Width           =   3225
      Begin VSFlex8Ctl.VSFlexGrid VSFType 
         Height          =   1995
         Left            =   30
         TabIndex        =   1
         Top             =   510
         Width           =   2895
         _cx             =   5106
         _cy             =   3519
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
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
         BackColorSel    =   16772055
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483635
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   2
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   3
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   350
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   -1  'True
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
         ShowComboButton =   0
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
      Begin VB.Label lblShortCaptionType 
         Caption         =   "未分组项目"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   0
         TabIndex        =   9
         Top             =   30
         Width           =   2535
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   6450
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   19394
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin XtremeSuiteControls.ShortcutCaption ShortCaptionType 
      Height          =   105
      Left            =   9420
      TabIndex        =   10
      Top             =   1260
      Width           =   2745
      _Version        =   589884
      _ExtentX        =   4842
      _ExtentY        =   185
      _StockProps     =   6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
      GradientColorLight=   14737632
      GradientColorDark=   14737632
   End
   Begin XtremeCommandBars.CommandBars cbrthis 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmAppforBillGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngkeyID As Long                   '分类id
Private mblnfrmIfShow As Boolean            '是否已加载
Private mblnEdit As Boolean                 '是否编辑数据
Private mblnItemSort As Boolean             '是否处于顺序调整状态

'实现拖动效果需要的变量
Private mlngMouseRow As Long                '新增的行
Private mlngMouseDownRow As Long            '鼠标按下的行

Private Sub cboGroupSel_Click()
    ReadItemData
End Sub

Private Sub cbrthis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case ConMenu_Appfro_AddBill     '增加
            Control.Enabled = cboGroupSel.ListCount < 20 And Not mblnItemSort
        Case ConMenu_Appfro_DelBill     '删除
            Control.Enabled = Not mblnItemSort
        Case ConMenu_Browse_Save
            Control.Enabled = mblnEdit
    End Select
    cboGroupSel.Enabled = Not mblnItemSort
End Sub

Private Sub cbrthis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case ConMenu_Appfro_AddBill                     '增加申请单
            frmAppforBillGroupItem.showMe Me, mlngkeyID, 0, "", ""
            LoadGroup
        Case ConMenu_Appfro_DelBill                     '删除申请单
            If cboGroupSel.ListCount >= 1 Then
                If Me.cboGroupSel.ItemData(Me.cboGroupSel.ListIndex) <> 0 Then
                    If MsgBox("您确定要删除该分组，删除分组后，对应的分组项目将同步清空，" & vbCrLf & "变成未分组项目。", vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
                        DelGroup Me.cboGroupSel.ItemData(Me.cboGroupSel.ListIndex)
                        LoadGroup
                    End If
                End If
            Else
                MsgBox "请先添加分组，在进行选择对应的项目！", vbInformation, "中联信息"
            End If
        Case ConMenu_Appfor_ItemSort                    '调整顺序
            mblnItemSort = True
            mblnEdit = True
        Case ConMenu_Browse_Save                        '保存
            SaveGroup
            ReadItemData
        Case ConMenu_Appfro_Exit                        '退出
            Unload Me
    End Select
End Sub

Private Sub DelGroup(ByVal lngGroupId As Long)
          Dim strSQL As String
              
          '保存
1         On Error GoTo DelGroup_Error

2         strSQL = "Zl_检验申请单分组_EDIT(2," & lngGroupId & ")"
3         ComExecuteProc Sel_Lis_DB, strSQL, "保存申请分类"
4         SaveDBLog 18, 6, 0, "删除", "删除项目分组:" & cboGroupSel.Text, 1012, "申请单设置"


5         Exit Sub
DelGroup_Error:
6         Call writeErrLog("zl9LisInsideComm", "frmAppforBillGroup", "执行(DelGroup)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
7         Err.Clear
          
End Sub

Private Sub SaveGroup()
          Dim strSQL As String
          Dim strID As String
          Dim i As Integer
          Dim strName As String
          Dim lngCount As Long
          
          '保存
1         On Error GoTo SaveGroup_Error

2         If cboGroupSel.ListCount >= 1 Then
3             If Me.cboGroupSel.ItemData(Me.cboGroupSel.ListIndex) <> 0 Then
4                 With vsfList
                      
5                     For i = .FixedRows To .Rows - 1
6                         strID = strID & "," & .TextMatrix(i, .ColIndex("id"))
7                         strName = strName & "," & .TextMatrix(i, .ColIndex("组合项目"))
8                     Next
                      
9                     If strID <> "" Then strID = Mid(strID, 2)
10                    If strName <> "" Then strName = Mid(strName, 2)
                      
11                    strSQL = "Zl_申请单明细分组_Update(" & Me.cboGroupSel.ItemData(Me.cboGroupSel.ListIndex) & "," & mlngkeyID & ",'" & strID & "')"
12                    ComExecuteProc Sel_Lis_DB, strSQL, "保存申请分组设置"
13                    SaveDBLog 18, 6, 0, "编辑", "编辑项目分组:" & cboGroupSel.Text & ",分组内容:" & strName, 1012, "申请单设置"
14                    If mblnItemSort Then
15                        For i = 1 To .Rows - 1
16                            If .RowHidden(i) = False Then
17                                lngCount = lngCount + 1
18                                strSQL = "Zl_申请单明细_Sort(" & Val(.TextMatrix(i, .ColIndex("明细ID"))) & "," & lngCount & ")"
19                                Call ComExecuteProc(Sel_Lis_DB, strSQL, "申请单排序")
20                            End If
21                        Next
22                    End If
23                End With
24            End If
25            mblnEdit = False
26            mblnItemSort = False
27        Else
28            MsgBox "请先添加分组，在进行选择对应的项目！", vbInformation, "中联信息"
29        End If
          

30        Exit Sub
SaveGroup_Error:
31        Call writeErrLog("zl9LisInsideComm", "frmAppforBillGroup", "执行(SaveGroup)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
32        Err.Clear
End Sub


Private Sub cbrthis_ResizeClient(ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)
    With ShortCaptionType
        .Top = Top - 10
        .Left = Left + 10
        .Width = (Right - Left)
    End With
    With Me.picLeft
        .Top = Top + 20
        .Left = Left + 10
        .Width = (Right - Left) * 2 / 5
        .Height = Bottom - Top - stbThis.Height + 1
    End With
    With Me.picRight
        .Top = Top + 20
        .Left = picLeft.Left + picLeft.Width + 1
        .Width = (Right - Left) - .Left - 25
        .Height = Me.picLeft.Height
    End With
    
End Sub



Private Sub picFindNext_Click()
          '添加分组
          Dim i As Integer
1         On Error GoTo picFindNext_Click_Error

2         If cboGroupSel.ListCount >= 1 Then
3             If Me.cboGroupSel.ItemData(Me.cboGroupSel.ListIndex) <> 0 Then
4                 If VSFType.TextMatrix(VSFType.Row, VSFType.ColIndex("组合项目")) <> "" Then
5                     With vsfList
6                         If VSFType.Row <> 0 Then
7                             .Rows = .Rows + 1
8                             .TextMatrix(.Rows - 1, .ColIndex("序号")) = .Rows - 1
9                             .TextMatrix(.Rows - 1, .ColIndex("id")) = VSFType.TextMatrix(VSFType.Row, VSFType.ColIndex("id"))
10                            .TextMatrix(.Rows - 1, .ColIndex("编码")) = VSFType.TextMatrix(VSFType.Row, VSFType.ColIndex("编码"))
11                            .TextMatrix(.Rows - 1, .ColIndex("组合项目")) = VSFType.TextMatrix(VSFType.Row, VSFType.ColIndex("组合项目"))
12                            .TextMatrix(.Rows - 1, .ColIndex("分组")) = Mid(cboGroupSel.Text, InStr(cboGroupSel.Text, "-") + 1)
13                        End If
14                    End With
15                    With VSFType
16                        If .Row <> 0 Then
17                            .RemoveItem .Row
18                            For i = .FixedRows To .Rows - 1
19                                .TextMatrix(i, .ColIndex("序号")) = i
20                            Next
21                        End If
22                    End With
23                End If
24            End If
25            If VSFType.Rows > 1 Then VSFType.Row = 1
26            If vsfList.Rows > 1 Then vsfList.Row = 1
27            mblnEdit = True
28        Else
29            MsgBox "请先添加分组，在进行选择对应的项目！", vbInformation, "中联信息"
30        End If


31        Exit Sub
picFindNext_Click_Error:
32        Call writeErrLog("zl9LisInsideComm", "frmAppforBillGroup", "执行(picFindNext_Click)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
33        Err.Clear
End Sub

Private Sub picLast_Click()
          '添加分组
          Dim i As Integer
1         On Error GoTo picLast_Click_Error

2         If cboGroupSel.ListCount >= 1 Then
3             If Me.cboGroupSel.ItemData(Me.cboGroupSel.ListIndex) <> 0 Then
4                 If vsfList.TextMatrix(vsfList.Row, vsfList.ColIndex("组合项目")) <> "" Then
5                     With VSFType
6                         If vsfList.Row <> 0 Then
7                             .Rows = .Rows + 1
8                             .TextMatrix(.Rows - 1, .ColIndex("序号")) = .Rows - 1
9                             .TextMatrix(.Rows - 1, .ColIndex("id")) = vsfList.TextMatrix(vsfList.Row, vsfList.ColIndex("id"))
10                            .TextMatrix(.Rows - 1, .ColIndex("编码")) = vsfList.TextMatrix(vsfList.Row, vsfList.ColIndex("编码"))
11                            .TextMatrix(.Rows - 1, .ColIndex("组合项目")) = vsfList.TextMatrix(vsfList.Row, vsfList.ColIndex("组合项目"))
12                        End If
13                    End With
14                    With vsfList
15                        If .Row <> 0 Then
16                            .RemoveItem .Row
17                            For i = .FixedRows To .Rows - 1
18                                .TextMatrix(i, .ColIndex("序号")) = i
19                            Next
20                        End If
21                    End With
22                End If
23            End If
24            If VSFType.Rows > 1 Then VSFType.Row = 1
25            If vsfList.Rows > 1 Then vsfList.Row = 1
26            mblnEdit = True
27        Else
28            MsgBox "请先添加分组，在进行选择对应的项目！", vbInformation, "中联信息"
29        End If


30        Exit Sub
picLast_Click_Error:
31        Call writeErrLog("zl9LisInsideComm", "frmAppforBillGroup", "执行(picLast_Click)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
32        Err.Clear
End Sub

Private Sub picLeft_Resize()
    On Error Resume Next
    With lblShortCaptionType
        .Top = 20
        .Left = 40
        .Width = Me.picLeft.ScaleWidth
    End With
    With VSFType
        .Top = lblShortCaptionType.Height + 10
        .Left = 40
        .Width = Me.picLeft.ScaleWidth - 40
        .Height = picLeft.ScaleHeight - .Top
    End With
End Sub

Private Sub picRight_Resize()
    On Error Resume Next
    With picFindNext
        .Top = picRight.Height / 2 - picLast.Height / 2 - 350
        .Left = 60
    End With
    With picLast
        .Top = picRight.Height / 2 + picFindNext.Height / 2 + 350
        .Left = 60
    End With

    With lblShortCaptionItem
        .Top = 20
        .Left = picFindNext.Left + picFindNext.Width + 80
        .Width = Me.picRight.ScaleWidth - picLast.Width - 160
    End With

    With vsfList
        .Top = lblShortCaptionItem.Height + 10
        .Left = picFindNext.Left + picFindNext.Width + 80
        .Width = Me.picRight.ScaleWidth - picLast.Width - 160
        .Height = picRight.ScaleHeight - .Top
    End With
End Sub

Private Sub picLast_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picLast.BorderStyle = 1
End Sub

Private Sub picLast_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picLast.BorderStyle = 0
End Sub

Private Sub picFindNext_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picFindNext.BorderStyle = 1
End Sub

Private Sub picFindNext_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picFindNext.BorderStyle = 0
End Sub

Public Sub showMe(frmObj As Object, lngkeyID As Long, strName As String)
    '功能       打开窗体并传入参数
    mlngkeyID = lngkeyID
    Me.Caption = "申请单下属分组设置(" & strName & ")"
    Me.Show vbModal, frmObj
    
End Sub

Private Sub Form_Load()
    '功能创建工具条
    Dim cbrControl As CommandBarControl
    Dim cbrToolBar As CommandBar
    Dim cbrCustom As CommandBarControlCustom
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbrthis.VisualTheme = xtpThemeOffice2003
    Me.cbrthis.Icons = frmPubIcons.imgPublic.Icons
    With Me.cbrthis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbrthis.EnableCustomization False

    '-----------------------------------------------------
    '菜单定义
    Me.cbrthis.ActiveMenuBar.Title = "菜单"
    Me.cbrthis.ActiveMenuBar.Visible = False
    Set cbrToolBar = Me.cbrthis.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, ConMenu_Appfro_AddBill, "添加分组")
        Set cbrControl = .Add(xtpControlButton, ConMenu_Appfro_DelBill, "删除分组")
        Set cbrControl = .Add(xtpControlButton, ConMenu_Appfor_ItemSort, "调整顺序")
        Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_Save, "保存")
        Set cbrControl = .Add(xtpControlButton, ConMenu_Appfro_Exit, "退出")
        cbrControl.BeginGroup = True
    End With
    
    Set cbrControl = cbrToolBar.Controls.Add(xtpControlLabel, 0, Space(10) & "请选择分组  ")
    cbrControl.Flags = xtpFlagRightAlign

    Set cbrCustom = cbrToolBar.Controls.Add(xtpControlCustom, ConMenu_Appfro_DeptSel, Space(10) & "     请选择分组  ")
    cbrCustom.ShortcutText = Space(10) & "     请选择分组  "
    cbrCustom.Handle = Me.cboGroupSel.hWnd
    cbrCustom.Flags = xtpFlagLeftPopup
    cbrCustom.Style = xtpButtonIconAndCaption
    For Each cbrControl In cbrToolBar.Controls
        If cbrControl.Type = xtpControlButton Then
            cbrControl.Style = xtpButtonIconAndCaption
        End If
    Next
End Sub

Private Sub Form_Activate()
    If mblnfrmIfShow = False Then
        Call InitVSF
        LoadGroup
        ReadItemData
        mblnfrmIfShow = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mblnfrmIfShow = False
    mblnEdit = False
    mblnItemSort = False
    mlngMouseRow = 0
    mlngMouseDownRow = 0
End Sub

Private Sub LoadGroup()
          '功能   读入分组
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          
1         On Error GoTo LoadGroup_Error

2         strSQL = "Select distinct ID, 编码, 名称 From 检验申请单分组 where 申请单id =[1] order by ID "
3         Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "读入申请单分组", mlngkeyID)
4         With cboGroupSel
5             .Clear
      '        .AddItem ""
      '        .ItemData(.NewIndex) = 0
6             Do Until rsTmp.EOF
7                 .AddItem rsTmp("编码") & "-" & rsTmp("名称")
8                 .ItemData(.NewIndex) = rsTmp("ID")
9                 rsTmp.MoveNext
10            Loop
11            If .ListCount > 0 Then .ListIndex = 0
12        End With


13        Exit Sub
LoadGroup_Error:
14        Call writeErrLog("zl9LisInsideComm", "frmAppforBillGroup", "执行(LoadGroup)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
15        Err.Clear
          
End Sub

Private Sub InitVSF()
          '初始化列表
          '左侧未选中
1         On Error GoTo InitVSF_Error

2         With Me.VSFType
3             .Rows = 2
4             .Cols = 4
5             .FixedRows = 1
6             .ColKey(0) = "ID": .ColWidth(.ColIndex("ID")) = 1000: .ColAlignment(.ColIndex("ID")) = flexAlignLeftCenter: .TextMatrix(0, .ColIndex("ID")) = "ID": .ColHidden(.ColIndex("ID")) = True
              
7             .ColKey(1) = "序号": .ColWidth(.ColIndex("序号")) = 600: .ColAlignment(.ColIndex("序号")) = flexAlignLeftCenter: .TextMatrix(0, .ColIndex("序号")) = "序号"
8                 .Cell(flexcpAlignment, 0, .ColIndex("序号"), 0, .ColIndex("序号")) = flexAlignCenterCenter
9              .ColKey(2) = "组合项目": .ColWidth(.ColIndex("组合项目")) = 3000: .ColAlignment(.ColIndex("组合项目")) = flexAlignLeftCenter: .TextMatrix(0, .ColIndex("组合项目")) = "组合项目"
10                .Cell(flexcpAlignment, 0, .ColIndex("组合项目"), 0, .ColIndex("组合项目")) = flexAlignCenterCenter
11            .ColKey(3) = "编码": .ColWidth(.ColIndex("编码")) = 900: .ColAlignment(.ColIndex("编码")) = flexAlignLeftCenter: .TextMatrix(0, .ColIndex("编码")) = "编码"
12                .Cell(flexcpAlignment, 0, .ColIndex("编码"), 0, .ColIndex("编码")) = flexAlignCenterCenter
13               End With
          '右侧已选中
14        With Me.vsfList
15            .Rows = 2
16            .Cols = 7
17            .FixedRows = 1
18            .ColKey(0) = "ID": .ColWidth(.ColIndex("ID")) = 1000: .ColAlignment(.ColIndex("ID")) = flexAlignLeftCenter: .TextMatrix(0, .ColIndex("ID")) = "ID": .ColHidden(.ColIndex("ID")) = True
19            .ColKey(1) = "序号": .ColWidth(.ColIndex("序号")) = 600: .ColAlignment(.ColIndex("序号")) = flexAlignLeftCenter: .TextMatrix(0, .ColIndex("序号")) = "序号"
20                .Cell(flexcpAlignment, 0, .ColIndex("序号"), 0, .ColIndex("序号")) = flexAlignCenterCenter
21            .ColKey(2) = "组合项目": .ColWidth(.ColIndex("组合项目")) = 3000: .ColAlignment(.ColIndex("组合项目")) = flexAlignLeftCenter: .TextMatrix(0, .ColIndex("组合项目")) = "组合项目"
22                .Cell(flexcpAlignment, 0, .ColIndex("组合项目"), 0, .ColIndex("组合项目")) = flexAlignCenterCenter
23            .ColKey(3) = "分组": .ColWidth(.ColIndex("分组")) = 1500: .ColAlignment(.ColIndex("分组")) = flexAlignLeftCenter: .TextMatrix(0, .ColIndex("分组")) = "分组"
24                .Cell(flexcpAlignment, 0, .ColIndex("分组"), 0, .ColIndex("分组")) = flexAlignCenterCenter
25            .ColKey(4) = "编码": .ColWidth(.ColIndex("编码")) = 1000: .ColAlignment(.ColIndex("编码")) = flexAlignLeftCenter: .TextMatrix(0, .ColIndex("编码")) = "编码"
26                .Cell(flexcpAlignment, 0, .ColIndex("编码"), 0, .ColIndex("编码")) = flexAlignCenterCenter
27            .ColKey(5) = "排列顺序": .ColHidden(.ColIndex("排列顺序")) = True: .ColAlignment(.ColIndex("排列顺序")) = flexAlignLeftCenter: .TextMatrix(0, .ColIndex("排列顺序")) = "排列顺序"
28                .Cell(flexcpAlignment, 0, .ColIndex("排列顺序"), 0, .ColIndex("排列顺序")) = flexAlignCenterCenter
29            .ColKey(6) = "明细ID": .ColHidden(.ColIndex("明细ID")) = True: .ColAlignment(.ColIndex("明细ID")) = flexAlignLeftCenter: .TextMatrix(0, .ColIndex("明细ID")) = "明细ID"
30                .Cell(flexcpAlignment, 0, .ColIndex("明细ID"), 0, .ColIndex("明细ID")) = flexAlignCenterCenter
                  
                  
31        End With
              


32        Exit Sub
InitVSF_Error:
33        Call writeErrLog("zl9LisInsideComm", "frmAppforBillGroup", "执行(InitVSF)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
34        Err.Clear
          
End Sub

Private Sub ReadItemData()
          '功能       读入分类明细
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          
1         On Error GoTo ReadItemData_Error

2         strSQL = "Select b.编码, b.名称,b.id" & vbNewLine & _
                   " From 检验申请单明细 A, 检验组合项目 B,检验申请单 c Where a.申请单id =c.id and A.组合id = B.Id and b.停用日期 is null and a.分组id  is null and a.申请单ID = [1] order by b.编码 "
3         Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "读入分类明细", mlngkeyID)
4         With Me.VSFType
5             .Rows = 1
              
6             Do Until rsTmp.EOF
7                 .Rows = .Rows + 1
8                 .TextMatrix(.Rows - 1, .ColIndex("序号")) = .Rows - 1
9                 .TextMatrix(.Rows - 1, .ColIndex("id")) = rsTmp("id") & ""
10                .TextMatrix(.Rows - 1, .ColIndex("编码")) = rsTmp("编码") & ""
11                .TextMatrix(.Rows - 1, .ColIndex("组合项目")) = rsTmp("名称") & ""
12                rsTmp.MoveNext
13            Loop
      '        If .Rows = 1 Then .Rows = 2
14        End With
15        If Me.cboGroupSel.ListCount >= 1 Then
16            strSQL = "Select b.编码, b.名称,d.名称 分组,b.id,a.排列顺序,a.id 明细ID" & vbNewLine & _
                       " From 检验申请单明细 A, 检验组合项目 B,检验申请单 c ,检验申请单分组 d Where a.申请单id =c.id and A.组合id = B.Id and  d.id= a.分组id and d.申请单id=a.申请单id and b.停用日期 is null and a.申请单ID = [1] and d.id=[2] order by a.排列顺序, b.编码 "
17            Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "读入分类明细", mlngkeyID, Me.cboGroupSel.ItemData(Me.cboGroupSel.ListIndex))
18            With Me.vsfList
19                .Rows = 1
20                Do Until rsTmp.EOF
21                    .Rows = .Rows + 1
22                    .TextMatrix(.Rows - 1, .ColIndex("序号")) = .Rows - 1
23                    .TextMatrix(.Rows - 1, .ColIndex("id")) = rsTmp("id") & ""
24                    .TextMatrix(.Rows - 1, .ColIndex("编码")) = rsTmp("编码") & ""
25                    .TextMatrix(.Rows - 1, .ColIndex("组合项目")) = rsTmp("名称") & ""
26                    .TextMatrix(.Rows - 1, .ColIndex("分组")) = rsTmp("分组") & ""
27                    .TextMatrix(.Rows - 1, .ColIndex("排列顺序")) = rsTmp("排列顺序") & ""
28                    .TextMatrix(.Rows - 1, .ColIndex("明细ID")) = rsTmp("明细ID") & ""
29                    rsTmp.MoveNext
30                Loop
31            End With
32        End If
33        If VSFType.Rows > 1 Then VSFType.Row = 1
34        If vsfList.Rows > 1 Then vsfList.Row = 1


35        Exit Sub
ReadItemData_Error:
36        Call writeErrLog("zl9LisInsideComm", "frmAppforBillGroup", "执行(ReadItemData)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
37        Err.Clear

End Sub

'==========以下代码功能为:将右侧VSF列表中的数据拖动到左侧VSF中=============
'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2017/6/14
'功    能:模拟拖动，点击列表时，将标签定位到点击的位置，方便跟随鼠标移动
'入    参:
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Private Sub VSFList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
1         On Error GoTo VSFList_MouseDown_Error

2         If Button <> 1 Then Exit Sub
3         If Not mblnItemSort Then Exit Sub
          
4         With Me.vsfList
5             If .MouseRow <= 0 Or .MouseCol < 0 Then Exit Sub
6             Me.lblShow.Caption = .TextMatrix(.MouseRow, .ColIndex("组合项目"))
7             Me.lblShow.Tag = .TextMatrix(.MouseRow, .ColIndex("id")) & "|" & .TextMatrix(.MouseRow, .ColIndex("编码")) & "|" & .TextMatrix(.MouseRow, .ColIndex("组合项目")) & "|" & .TextMatrix(.MouseRow, .ColIndex("明细ID"))
8             mlngMouseDownRow = .MouseRow
9         End With


10        Exit Sub
VSFList_MouseDown_Error:
11        MsgBox "执行(VSFList_MouseDown)时出错,错误描述:" & Err.Description & " 错误号:" & Err.Number & " 错误行:" & Erl, vbInformation, "提示"
12        Err.Clear
End Sub

'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2017/6/14
'功    能:标签跟随鼠标移动
'入    参:
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Private Sub VSFList_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
          Dim lngRow As Long
          Dim lngCol As Long
          
1         On Error GoTo VSFList_MouseMove_Error

2         If Button <> 1 Then Exit Sub
3         If Not mblnItemSort Then Exit Sub
          
4         If Me.lblShow.Caption = "" Then Exit Sub
5         With Me.lblShow
6             If .Visible = False Then .Visible = True
7             .Left = X - (.Width / 2)
8             .Top = Y - (.Height / 2)
9         End With
          
          '设置右侧列表在拖动鼠标时的效果
10        With Me.vsfList
11            lngRow = .MouseRow
12            lngCol = .MouseCol
13            If lngRow > -1 And lngCol > -1 Then
14                If mlngMouseRow <> lngRow And mlngMouseRow > 0 And lngRow > 0 Then
                      '移动到某一行上之后新增一个空行
15                    If mlngMouseRow <= .Rows - 1 Then
16                        If Trim(.TextMatrix(mlngMouseRow, .ColIndex("组合项目"))) = "" Then .RemoveItem mlngMouseRow    '先移除之前的空行
17                    End If
18                    Debug.Print 1
19                    .AddItem "", lngRow
20                    mlngMouseRow = lngRow
21                    .Row = mlngMouseRow
22                ElseIf mlngMouseRow = 0 And lngRow > 0 Then
23                    Debug.Print 2
24                    .AddItem "", lngRow
25                    mlngMouseRow = lngRow
26                ElseIf lngRow = .Rows - 1 And Trim(.TextMatrix(.Rows - 1, .ColIndex("组合项目"))) <> "" Then
                      '如果移动到最后一行,则在最后新增一行
27                    Debug.Print 3
28                    .AddItem "", .Rows
29                    mlngMouseRow = .Rows
30                End If
31            ElseIf lngRow = -1 And .Rows < 2 Then
32                Debug.Print 4
33                .Rows = .Rows + 1
34                mlngMouseRow = .Rows - 1
35            ElseIf lngRow = -1 And lngCol = -1 And mlngMouseRow <= .Rows - 1 Then
36                If Trim(.TextMatrix(mlngMouseRow, .ColIndex("组合项目"))) = "" Then
37                    .RemoveItem mlngMouseRow
38                End If
39            End If
40        End With
          

41        Exit Sub
VSFList_MouseMove_Error:
42        MsgBox "执行(VSFList_MouseMove)时出错,错误描述:" & Err.Description & " 错误号:" & Err.Number & " 错误行:" & Erl, vbInformation, "提示"
43        Err.Clear

End Sub


'---------------------------------------------------------------------------------------
'编    码:蔡青松
'编码时间:2017/6/14
'功    能:松开鼠标时,将拖动的值复制到右边的VSF中
'入    参:
'出    参:
'返    回:
'---------------------------------------------------------------------------------------
Private Sub VSFList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    
    On Error GoTo VSFList_MouseUp_Error

    If Button <> 1 Then Exit Sub
    If Not mblnItemSort Then Exit Sub
    With Me.vsfList
        '仅在右侧列表中拖动数据时
        If .MouseCol > -1 And mlngMouseRow > 0 And mlngMouseRow <= .Rows - 1 Then
            .TextMatrix(mlngMouseRow, .ColIndex("id")) = Split(Me.lblShow.Tag, "|")(0): .ColAlignment(.ColIndex("id")) = flexAlignLeftCenter
            .TextMatrix(mlngMouseRow, .ColIndex("编码")) = Split(Me.lblShow.Tag, "|")(1): .ColAlignment(.ColIndex("编码")) = flexAlignLeftCenter
            .TextMatrix(mlngMouseRow, .ColIndex("组合项目")) = Split(Me.lblShow.Tag, "|")(2): .ColAlignment(.ColIndex("组合项目")) = flexAlignLeftCenter
            .TextMatrix(mlngMouseRow, .ColIndex("明细ID")) = Split(Me.lblShow.Tag, "|")(3): .ColAlignment(.ColIndex("明细ID")) = flexAlignLeftCenter
            If mlngMouseDownRow > 0 And Me.lblShow.Visible = True Then
                If mlngMouseRow > mlngMouseDownRow Then
                    .RemoveItem mlngMouseDownRow
                ElseIf mlngMouseDownRow + 1 <= .Rows - 1 Then
                    If .MouseCol > -1 Then .RemoveItem mlngMouseDownRow + 1
                End If
            End If
        End If
        
        For i = 1 To .Rows - 1
            .TextMatrix(i, .ColIndex("序号")) = i
        Next
        
    End With
    mlngMouseRow = 0
    Me.lblShow.Caption = ""
    If Me.lblShow.Visible = True Then Me.lblShow.Visible = False
    
    

    Exit Sub
VSFList_MouseUp_Error:
    MsgBox "执行(VSFList_MouseUp)时出错,错误描述:" & Err.Description & " 错误号:" & Err.Number & " 错误行:" & Erl, vbInformation, "提示"
    'WriteLog "执行(VSFList_MouseUp)时出错,错误描述:" & Err.Description & " 错误号:" & Err.Number & " 错误行:" & Erl
    Err.Clear
End Sub
'=============================================================

