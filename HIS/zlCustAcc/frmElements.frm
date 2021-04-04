VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmElements 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "记帐单元素"
   ClientHeight    =   5955
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   9525
   Icon            =   "frmElements.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   9525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox txt数次 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   2700
      TabIndex        =   10
      Top             =   1500
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.TextBox txtEdit 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   3120
      TabIndex        =   11
      Top             =   1170
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.ComboBox cmb类别 
      Height          =   300
      Left            =   3090
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   825
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "…"
      Height          =   300
      Left            =   4500
      TabIndex        =   12
      Top             =   1140
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Frame fra可选 
      Caption         =   "可选项目"
      Height          =   3285
      Left            =   120
      TabIndex        =   6
      Top             =   1590
      Width           =   2745
      Begin VB.ListBox lst项目 
         Height          =   2790
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   7
         Top             =   270
         Width           =   2475
      End
   End
   Begin VB.Frame fraFix 
      Caption         =   "固定项目"
      Height          =   1275
      Left            =   120
      TabIndex        =   0
      Top             =   210
      Width           =   2745
      Begin VB.ComboBox cmb开单 
         Height          =   300
         Left            =   1170
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   750
         Width           =   1485
      End
      Begin MSComCtl2.UpDown ud项目数 
         Height          =   300
         Left            =   2385
         TabIndex        =   3
         Top             =   360
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Value           =   1
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txt项目数"
         BuddyDispid     =   196618
         OrigLeft        =   2370
         OrigTop         =   450
         OrigRight       =   2610
         OrigBottom      =   765
         Max             =   100
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txt项目数 
         Height          =   300
         Left            =   1500
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   360
         Width           =   885
      End
      Begin VB.Label lbl科室 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "开单科室(&D)"
         Height          =   180
         Left            =   120
         TabIndex        =   4
         Top             =   810
         Width           =   990
      End
      Begin VB.Label lbl项目数量 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "收费项目个数(&N)"
         Height          =   180
         Left            =   120
         TabIndex        =   1
         Top             =   420
         Width           =   1350
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6570
      TabIndex        =   15
      Tag             =   "分类"
      Top             =   5010
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   5250
      TabIndex        =   14
      Tag             =   "分类"
      Top             =   5010
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   360
      TabIndex        =   16
      Tag             =   "分类"
      Top             =   5010
      Width           =   1100
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msh收费项目 
      Height          =   4395
      Left            =   2970
      TabIndex        =   13
      Top             =   480
      Width           =   4965
      _ExtentX        =   8758
      _ExtentY        =   7752
      _Version        =   393216
      Rows            =   12
      RowHeightMin    =   320
      BackColorSel    =   -2147483643
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      GridColor       =   8421504
      GridColorFixed  =   8421504
      GridLinesFixed  =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label lbl收费项目 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "收费项目(&S)"
      Height          =   180
      Left            =   3030
      TabIndex        =   8
      Top             =   210
      Width           =   990
   End
End
Attribute VB_Name = "frmElements"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum RowSign
    row收费类别值 = 1
    row收费细目值 = 2
    row数次值 = 3
    row收费类别 = 4
    row细目选择 = 5
    row计算单位 = 6
    row数次 = 7
    row标准单价 = 8
    row应收金额 = 9
    row实收金额 = 10
    row执行部门 = 11
    row附加标志 = 12
End Enum

Dim mblnNew As Boolean  '当前修改的单据是否是新增加
Dim mcolBill As Elements
Dim mblnOK As Boolean
Dim mblnChange As Boolean
Dim mlngItemCount As Long

Public Function ModifyElement(colBill As Elements, ItemCount As Long, Optional ByVal 新增 As Boolean) As Boolean
'修改“记帐单元素”的管理程序
    Dim lngCount As Long, lngRow As Long
    Dim lngID As Long, strTemp As String
    
    mblnOK = False
    mblnNew = 新增
    mlngItemCount = ItemCount
    Set mcolBill = colBill
    
    
    Call InitCont
    If 新增 = True Then
        cmb开单.ListIndex = 0
        For lngCount = 0 To lst项目.ListCount - 1
            lst项目.Selected(lngCount) = True
        Next
        
        msh收费项目.TextMatrix(row数次值, 1) = "1"
        For lngCount = row收费类别 To msh收费项目.Rows - 1
            msh收费项目.TextMatrix(lngCount, 1) = "√"
        Next
    Else
        '固定项目
        ud项目数.Value = ItemCount '自动会改变表格的列数
        
        If Mid(colBill("开单部门").Value, 1, 1) = "C" Then
            cmb开单.ListIndex = Mid(colBill("开单部门").Value, 2, 1)
        ElseIf colBill("开单部门").Value <> "" Then
            lngID = Val(colBill("开单部门").Value)
            For lngCount = 4 To cmb开单.ListCount - 1
                If cmb开单.ItemData(lngCount) = lngID Then
                    cmb开单.ListIndex = lngCount
                    Exit For
                End If
            Next
            If cmb开单.ListIndex < 0 Then cmb开单.ListIndex = 0
        Else
            cmb开单.ListIndex = 0
        End If
        
        '可选项目
        For lngCount = 0 To lst项目.ListCount - 1
            lst项目.Selected(lngCount) = colBill(lst项目.List(lngCount)).Visible
        Next
        '收费项目
        With msh收费项目
            For lngCount = 1 To ud项目数.Value
                For lngRow = 1 To .Rows - 1
                    Select Case .TextMatrix(lngRow, 0)
                        Case "收费类别值"
                            .TextMatrix(lngRow, lngCount) = GetClassName(colBill("收费类别" & "_" & lngCount).Value)
                        Case "收费细目值"
                            lngID = colBill("收费细目" & "_" & lngCount).Value
                            strTemp = GetItemName(Abs(lngID))
                            If strTemp <> "" Then
                                .ColData(lngCount) = lngID
                                .TextMatrix(lngRow, lngCount) = strTemp
                            End If
                        Case "数次值"
                            .TextMatrix(lngRow, lngCount) = colBill("数次" & "_" & lngCount).Value
                        Case Else
                            .TextMatrix(lngRow, lngCount) = IIf(colBill(.TextMatrix(lngRow, 0) & "_" & lngCount).Visible, "√", "")
                    End Select
                Next
            Next
            .Row = row收费类别
            .LeftCol = 1
            
        End With
        
    End If
    
    txt项目数.Text = ud项目数.Value
    cmb类别.Visible = False
    mblnChange = False
    frmElements.Show vbModal, frmDesign
    ModifyElement = mblnOK
    '返回行数
    If mblnOK = True Then
        ItemCount = mlngItemCount
    End If
End Function

Private Sub cmb开单_Click()
    '该项只是作为分隔符使用
    mblnChange = True
    If cmb开单.Text = "──────────────" Then cmb开单.ListIndex = 0
End Sub

Private Sub cmb类别_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Dim blnCancel As Boolean
        
        Call cmb类别_Validate(blnCancel)
        msh收费项目.Row = 2
        Call msh收费项目_EnterCell
    End If
End Sub

Private Sub cmb类别_Validate(Cancel As Boolean)
    If msh收费项目.Text <> cmb类别.Text Then
        msh收费项目.Text = cmb类别.Text
        msh收费项目.TextMatrix(row收费细目值, msh收费项目.Col) = ""
        msh收费项目.ColData(msh收费项目.Col) = 0
        If Left(cmb类别.Text, 1) = "0" Then
            msh收费项目.TextMatrix(row收费类别, msh收费项目.Col) = "√"
        End If
    End If
    
    mblnChange = True
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
ShowHelp "zl9custacc", Me.hwnd, Me.Name
End Sub

Private Sub cmdOK_Click()
    cmb类别.Visible = False
    Call SaveControls
    mblnOK = True
    mblnChange = False
    mlngItemCount = ud项目数.Value
    Unload Me
End Sub

Private Sub SaveControls()
    Dim ctlTemp As Control
    Dim lngCount As Long
    Dim lngRow As Long
    Dim blnVisible As Boolean
    Dim strTemp As String, strControl As String
    Dim lngLeft As Long, lngWidth As Long
    Dim arrItems As Variant
    
    If mblnNew = True Then
        '首先把那些没出现在对话框中，但又是必须的控件加上
        mcolBill.Clear
        '标题
            Set ctlTemp = LoadControl("Label", 225, 180, , , 0)
            mcolBill.Add "标签_0", ctlTemp, , True
        'No
            Set ctlTemp = LoadControl("Label", 8685, 720)
            mcolBill.Add "标签_" & ctlTemp.Index, ctlTemp, , True
            ctlTemp.Caption = "单据号"
            Set ctlTemp = LoadControl("ComboBox", 9465, 660, 1425, , 0)
            mcolBill.Add "NO", ctlTemp, , True
        '发生时间
            Set ctlTemp = LoadControl("Label", 8160, 5580)
            mcolBill.Add "标签_" & ctlTemp.Index, ctlTemp, , True
            ctlTemp.Caption = "时间"
            Set ctlTemp = LoadControl("TextBox", 8610, 5520, 2400, , 0)
            mcolBill.Add "发生时间", ctlTemp, , True
        '姓名
            Set ctlTemp = LoadControl("Label", 75, 1125)
            mcolBill.Add "标签_" & ctlTemp.Index, ctlTemp, , True
            ctlTemp.Caption = "姓名"
            Set ctlTemp = LoadControl("TextBox", 555, 1065, 1365)
            mcolBill.Add "姓名", ctlTemp, , True
        '开单部门id
            Set ctlTemp = LoadControl("Label", 8760, 1140)
            mcolBill.Add "标签_" & ctlTemp.Index, ctlTemp, , True
            ctlTemp.Caption = "科室"
            Set ctlTemp = LoadControl("ComboBox", 9195, 1080, 2055)
            mcolBill.Add "开单部门", ctlTemp, , True
        '开单人
            Set ctlTemp = LoadControl("Label", 5250, 5585)
            mcolBill.Add "标签_" & ctlTemp.Index, ctlTemp, , True
            ctlTemp.Caption = "医生"
            Set ctlTemp = LoadControl("ComboBox", 5625, 5520, 2085)
            mcolBill.Add "开单人", ctlTemp, , True
        '确定
            Set ctlTemp = LoadControl("CommandButton", 8295, 6120, , , 1)
            mcolBill.Add "确定", ctlTemp, , True
        '取消
            Set ctlTemp = LoadControl("CommandButton", 9690, 6120, , , 2)
            mcolBill.Add "取消", ctlTemp, , True
        '销
            Set ctlTemp = LoadControl("CheckBox", 10890, 660, 400, , 1)
            mcolBill.Add "销", ctlTemp, , True
            
        '然后处理可选取项目
        For lngCount = 0 To lst项目.ListCount - 1
            blnVisible = lst项目.Selected(lngCount)
            Select Case lst项目.List(lngCount)
               Case "病人ID"
                    If blnVisible = True Then
                        Set ctlTemp = LoadControl("Label", 75, 1510)
                        mcolBill.Add "标签_" & ctlTemp.Index, ctlTemp, , True
                        ctlTemp.Caption = "病人ID"
                    End If
                    Set ctlTemp = LoadControl("TextBox", 685, 1455, 1005)
                    mcolBill.Add "病人ID", ctlTemp, , blnVisible
               Case "标识号"
                    If blnVisible = True Then
                        Set ctlTemp = LoadControl("Label", 1875, 1510)
                        mcolBill.Add "标签_" & ctlTemp.Index, ctlTemp, , True
                        ctlTemp.Caption = "标识号"
                    End If
                    Set ctlTemp = LoadControl("TextBox", 2475, 1455, 1005)
                    mcolBill.Add "标识号", ctlTemp, , blnVisible
               Case "入院次数"
                    If blnVisible = True Then
                        Set ctlTemp = LoadControl("Label", 3585, 1510)
                        mcolBill.Add "标签_" & ctlTemp.Index, ctlTemp, , True
                        ctlTemp.Caption = "入院次数"
                    End If
                    Set ctlTemp = LoadControl("TextBox", 4350, 1455, 1005)
                    mcolBill.Add "入院次数", ctlTemp, , blnVisible
               Case "性别"
                    If blnVisible = True Then
                        Set ctlTemp = LoadControl("Label", 2085, 1125)
                        mcolBill.Add "标签_" & ctlTemp.Index, ctlTemp, , True
                        ctlTemp.Caption = "性别"
                    End If
                    Set ctlTemp = LoadControl("ComboBox", 2475, 1065, 1005)
                    mcolBill.Add "性别", ctlTemp, , blnVisible
               Case "年龄"
                    If blnVisible = True Then
                        Set ctlTemp = LoadControl("Label", 3615, 1125)
                        mcolBill.Add "标签_" & ctlTemp.Index, ctlTemp, , True
                        ctlTemp.Caption = "年龄"
                    End If
                    Set ctlTemp = LoadControl("TextBox", 3990, 1065, 660)
                    mcolBill.Add "年龄", ctlTemp, , blnVisible
               Case "床号"
                    If blnVisible = True Then
                        Set ctlTemp = LoadControl("Label", 6765, 1125)
                        mcolBill.Add "标签_" & ctlTemp.Index, ctlTemp, , True
                        ctlTemp.Caption = "床号"
                    End If
                    Set ctlTemp = LoadControl("TextBox", 7185, 1065, 1005)
                    mcolBill.Add "床号", ctlTemp, , blnVisible
               Case "病人病区"
                    If blnVisible = True Then
                        Set ctlTemp = LoadControl("Label", 5775, 1510)
                        mcolBill.Add "标签_" & ctlTemp.Index, ctlTemp, , True
                        ctlTemp.Caption = "病人病区"
                    End If
                    Set ctlTemp = LoadControl("TextBox", 6570, 1455, 1500)
                    mcolBill.Add "病人病区", ctlTemp, , blnVisible
               Case "病人科室"
                    If blnVisible = True Then
                        Set ctlTemp = LoadControl("Label", 8265, 1510)
                        mcolBill.Add "标签_" & ctlTemp.Index, ctlTemp, , True
                        ctlTemp.Caption = "病人科室"
                    End If
                    Set ctlTemp = LoadControl("TextBox", 9030, 1455, 1500)
                    mcolBill.Add "病人科室", ctlTemp, , blnVisible
               Case "费别"
                    If blnVisible = True Then
                        Set ctlTemp = LoadControl("Label", 4800, 1125)
                        mcolBill.Add "标签_" & ctlTemp.Index, ctlTemp, , True
                        ctlTemp.Caption = "费别"
                    End If
                    Set ctlTemp = LoadControl("ComboBox", 5190, 1065)
                    mcolBill.Add "费别", ctlTemp, , blnVisible
               Case "加班标志"
                    Set ctlTemp = LoadControl("CheckBox", 350, 5630, 800, , 0)
                    ctlTemp.Caption = "加班"
                    mcolBill.Add "加班标志", ctlTemp, , blnVisible
               Case "婴儿费"
                    If blnVisible = True Then
                        Set ctlTemp = LoadControl("Label", 1220, 5630)
                        mcolBill.Add "标签_" & ctlTemp.Index, ctlTemp, , True
                        ctlTemp.Caption = "婴儿费"
                    End If
                    Set ctlTemp = LoadControl("ComboBox", 1950, 5565, 1435)
                    mcolBill.Add "婴儿费", ctlTemp, , blnVisible
               Case "应收合计"
                    If blnVisible = True Then
                        Set ctlTemp = LoadControl("Label", 330, 6150)
                        mcolBill.Add "标签_" & ctlTemp.Index, ctlTemp, , True
                        ctlTemp.Caption = "应收"
                    End If
                    Set ctlTemp = LoadControl("TextBox", 720, 6105, 2085)
                    mcolBill.Add "应收合计", ctlTemp, , blnVisible
               Case "实收合计"
                    If blnVisible = True Then
                        Set ctlTemp = LoadControl("Label", 3000, 6150)
                        mcolBill.Add "标签_" & ctlTemp.Index, ctlTemp, , True
                        ctlTemp.Caption = "实收"
                    End If
                    Set ctlTemp = LoadControl("TextBox", 3390, 6105, 2085)
                    mcolBill.Add "实收合计", ctlTemp, , blnVisible
            End Select
        Next
        '收费项目
        With msh收费项目
            For lngCount = 1 To ud项目数.Value
                '从第2行开始
                For lngRow = 2 To .Rows - 1
                    strTemp = .TextMatrix(lngRow, 0)
                    blnVisible = .TextMatrix(lngRow, lngCount) = "√"
                    If lngCount = 1 Then
                        If blnVisible = True Or strTemp = "收费细目值" Then
                            Set ctlTemp = LoadControl("Label", 405, 2100)
                            mcolBill.Add "标签_" & ctlTemp.Index, ctlTemp, , True
                            ctlTemp.Caption = Replace(.TextMatrix(lngRow, 0), "值", "")
                        End If
                    End If
                    
                    
                    Select Case strTemp
                        Case "收费类别"
                            lngLeft = 120
                            lngWidth = 795
                            strControl = "ComboBox"
                        Case "收费细目值"
                            lngLeft = 915
                            lngWidth = 2115
                            strTemp = "收费细目"
                            strControl = "TextBox"
                            blnVisible = True '收费项目是肯定要显示的
                        Case "细目选择"
                            lngLeft = 3030
                            lngWidth = frmDesign.cmd(0).Width
                            strControl = "CommandButton"
                        Case "计算单位"
                            lngLeft = 3435
                            lngWidth = 1035
                            strControl = "TextBox"
                        Case "数次"
                            lngLeft = 4470
                            lngWidth = 705
                            strControl = "TextBox"
                        Case "标准单价"
                            lngLeft = 5175
                            lngWidth = 915
                            strControl = "TextBox"
                        Case "应收金额"
                            lngLeft = 6090
                            lngWidth = 1425
                            strControl = "TextBox"
                        Case "实收金额"
                            lngLeft = 7515
                            lngWidth = 1185
                            strControl = "TextBox"
                        Case "执行部门"
                            lngLeft = 8700
                            lngWidth = 1485
                            strControl = "ComboBox"
                        Case "附加标志"
                            lngLeft = 10215
                            lngWidth = 1065
                            strControl = "CheckBox"
                    End Select
                    
                    If strTemp <> "数次值" Then
                        '添加控件及元素
                        If lngCount = 1 Then
                            If blnVisible = True Then
                                ctlTemp.Left = lngLeft '标签的左边距
                            End If
                            
                            If strTemp = "细目选择" Then
                                Set ctlTemp = LoadControl(strControl, lngLeft, 2430, , , 0)
                            Else
                                Set ctlTemp = LoadControl(strControl, lngLeft, 2430, lngWidth)
                            End If
                        Else
                            Set ctlTemp = LoadControl(strControl, lngLeft, _
                                mcolBill(strTemp & "_" & lngCount - 1).Control.Top + mcolBill(strTemp & "_" & lngCount - 1).Control.Height, _
                                mcolBill(strTemp & "_" & lngCount - 1).Control.Width)
                        End If
                        
                        If strTemp = "附加标志" Then
                            ctlTemp.Caption = "附加手术"
                        End If
                        
                        mcolBill.Add strTemp & "_" & lngCount, ctlTemp, , blnVisible
                    End If
                    '预设值
                    If strTemp = "收费类别" Then
                        mcolBill(strTemp & "_" & lngCount).Value = Mid(.TextMatrix(1, lngCount), 1, 1)
                    ElseIf strTemp = "收费细目" Then
                        mcolBill(strTemp & "_" & lngCount).Value = .ColData(lngCount)
                    ElseIf strTemp = "数次" Then
                        mcolBill(strTemp & "_" & lngCount).Value = .TextMatrix(row数次值, lngCount)
                    End If
                Next
            Next
        End With
        
    Else
        '对现有控件进行修改
        
        '处理可选取项目
        For lngCount = 0 To lst项目.ListCount - 1
            mcolBill(lst项目.List(lngCount)).Visible = lst项目.Selected(lngCount)
        Next
        '处理收费细目
        
        arrItems = Split("收费类别,收费细目,细目选择,计算单位,数次,标准单价,应收金额,实收金额,执行部门,附加标志", ",")
        With msh收费项目
            If mlngItemCount > ud项目数.Value Then
                '现在比以前减少了，要删除一些
                For lngCount = ud项目数.Value + 1 To mlngItemCount
                    For lngRow = LBound(arrItems) To UBound(arrItems)
                        strTemp = arrItems(lngRow) & "_" & lngCount
                        '在删除之前先要卸装控件
                        Set ctlTemp = mcolBill(strTemp).Control
                        Unload ctlTemp
                        mcolBill.Remove strTemp
                    Next
                Next
            Else
                For lngCount = mlngItemCount + 1 To ud项目数.Value
                    For lngRow = LBound(arrItems) To UBound(arrItems)
                        strTemp = arrItems(lngRow)
                        Select Case strTemp
                            Case "收费类别"
                                blnVisible = .TextMatrix(row收费类别, lngCount) = "√"
                                Set ctlTemp = LoadControl("ComboBox", mcolBill(strTemp & "_" & lngCount - 1).Control.Left, mcolBill(strTemp & "_" & lngCount - 1).Control.Top + 300, 795)
                                mcolBill.Add strTemp & "_" & lngCount, ctlTemp, , blnVisible
                                mcolBill(strTemp & "_" & lngCount).Value = Mid(.TextMatrix(row收费类别值, lngCount), 1, 1)
                            Case "收费细目"
                                Set ctlTemp = LoadControl("TextBox", mcolBill(strTemp & "_" & lngCount - 1).Control.Left, mcolBill(strTemp & "_" & lngCount - 1).Control.Top + 300, 2115)
                                mcolBill.Add strTemp & "_" & lngCount, ctlTemp
                                mcolBill(strTemp & "_" & lngCount).Value = Mid(.TextMatrix(1, lngRow), 1, 1)
                                mcolBill(strTemp & "_" & lngCount).Value = .ColData(lngCount)
                            Case "细目选择"
                                blnVisible = .TextMatrix(row细目选择, lngCount) = "√"
                                Set ctlTemp = LoadControl("CommandButton", mcolBill(strTemp & "_" & lngCount - 1).Control.Left, mcolBill("细目选择" & "_" & lngCount - 1).Control.Top + 300)
                                mcolBill.Add strTemp & "_" & lngCount, ctlTemp, , blnVisible
                            Case "计算单位"
                                blnVisible = .TextMatrix(row计算单位, lngCount) = "√"
                                Set ctlTemp = LoadControl("TextBox", mcolBill(strTemp & "_" & lngCount - 1).Control.Left, mcolBill(strTemp & "_" & lngCount - 1).Control.Top + 300, 1035)
                                mcolBill.Add strTemp & "_" & lngCount, ctlTemp, , blnVisible
                            Case "数次"
                                blnVisible = .TextMatrix(row数次, lngCount) = "√"
                                Set ctlTemp = LoadControl("TextBox", mcolBill(strTemp & "_" & lngCount - 1).Control.Left, mcolBill(strTemp & "_" & lngCount - 1).Control.Top + 300, 705)
                                mcolBill.Add strTemp & "_" & lngCount, ctlTemp, , blnVisible
                                mcolBill(strTemp & "_" & lngCount).Value = .TextMatrix(row数次值, lngCount)
                            Case "标准单价"
                                blnVisible = .TextMatrix(row标准单价, lngCount) = "√"
                                Set ctlTemp = LoadControl("TextBox", mcolBill(strTemp & "_" & lngCount - 1).Control.Left, mcolBill(strTemp & "_" & lngCount - 1).Control.Top + 300, 915)
                                mcolBill.Add strTemp & "_" & lngCount, ctlTemp, , blnVisible
                            Case "应收金额"
                                blnVisible = .TextMatrix(row应收金额, lngCount) = "√"
                                Set ctlTemp = LoadControl("TextBox", mcolBill(strTemp & "_" & lngCount - 1).Control.Left, mcolBill(strTemp & "_" & lngCount - 1).Control.Top + 300, 1425)
                                mcolBill.Add strTemp & "_" & lngCount, ctlTemp, , blnVisible
                            Case "实收金额"
                                blnVisible = .TextMatrix(row实收金额, lngCount) = "√"
                                Set ctlTemp = LoadControl("TextBox", mcolBill(strTemp & "_" & lngCount - 1).Control.Left, mcolBill(strTemp & "_" & lngCount - 1).Control.Top + 300, 1185)
                                mcolBill.Add strTemp & "_" & lngCount, ctlTemp, , blnVisible
                            Case "执行部门"
                                blnVisible = .TextMatrix(row执行部门, lngCount) = "√"
                                Set ctlTemp = LoadControl("ComboBox", mcolBill(strTemp & "_" & lngCount - 1).Control.Left, mcolBill(strTemp & "_" & lngCount - 1).Control.Top + 300, 1485)
                                mcolBill.Add strTemp & "_" & lngCount, ctlTemp, , blnVisible
                            Case "附加标志"
                                blnVisible = .TextMatrix(row附加标志, lngCount) = "√"
                                Set ctlTemp = LoadControl("CheckBox", mcolBill(strTemp & "_" & lngCount - 1).Control.Left, mcolBill(strTemp & "_" & lngCount - 1).Control.Top + 300)
                                ctlTemp.Caption = "附加手术"
                                mcolBill.Add strTemp & "_" & lngCount, ctlTemp, , blnVisible
                        End Select
                    Next
                Next
            End If
            '更新其它的
            For lngCount = 1 To IIf(mlngItemCount < ud项目数.Value, mlngItemCount, ud项目数.Value) '取两者小的
                For lngRow = LBound(arrItems) To UBound(arrItems)
                    strTemp = arrItems(lngRow)
                    Select Case strTemp
                        Case "收费类别"
                            blnVisible = .TextMatrix(row收费类别, lngCount) = "√"
                            mcolBill(strTemp & "_" & lngCount).Visible = blnVisible
                            mcolBill(strTemp & "_" & lngCount).Value = Mid(.TextMatrix(row收费类别值, lngCount), 1, 1)
                        Case "收费细目"
                            mcolBill(strTemp & "_" & lngCount).Value = .ColData(lngCount)
                        Case Else
                            blnVisible = .TextMatrix(lngRow + 3, lngCount) = "√"
                            mcolBill(strTemp & "_" & lngCount).Visible = blnVisible
                            
                            If strTemp = "数次" Then
                                mcolBill(strTemp & "_" & lngCount).Value = .TextMatrix(row数次值, lngCount)
                            End If
                    End Select
                Next
            Next
        End With
    End If
    
    '保存开单科室的值
    If cmb开单.ListIndex >= 0 And cmb开单.ListIndex < 3 Then
        If cmb开单.ListIndex = 0 Or cmb开单.ListIndex = 3 Then
            mcolBill("开单部门").Value = ""
        Else
            mcolBill("开单部门").Value = "C" & cmb开单.ListIndex
        End If
    Else
        mcolBill("开单部门").Value = cmb开单.ItemData(cmb开单.ListIndex)
    End If
    
End Sub

Private Function LoadControl(ByVal ControlType As String, ByVal Left As Single, ByVal Top As Single, _
    Optional ByVal Width As Single, Optional ByVal Height As Single, Optional ByVal Index As Long = -1) As Control
    
    Dim ctl As Control
    '装载控件
    Select Case ControlType
        Case "ComboBox"
            If Index = -1 Then
               Load frmDesign.cmb(frmDesign.cmb.UBound + 1)
                Set ctl = frmDesign.cmb(frmDesign.cmb.UBound)
            Else
                Set ctl = frmDesign.cmb(Index)
            End If
        Case "CommandButton"
            If Index = -1 Then
               Load frmDesign.cmd(frmDesign.cmd.UBound + 1)
                Set ctl = frmDesign.cmd(frmDesign.cmd.UBound)
            Else
                Set ctl = frmDesign.cmd(Index)
            End If
        Case "CheckBox"
            If Index = -1 Then
               Load frmDesign.chk(frmDesign.chk.UBound + 1)
                Set ctl = frmDesign.chk(frmDesign.chk.UBound)
            Else
                Set ctl = frmDesign.chk(Index)
            End If
        Case "Label"
            If Index = -1 Then
               Load frmDesign.lbl(frmDesign.lbl.UBound + 1)
                Set ctl = frmDesign.lbl(frmDesign.lbl.UBound)
            Else
                Set ctl = frmDesign.lbl(Index)
            End If
        Case "TextBox"
            If Index = -1 Then
               Load frmDesign.txt(frmDesign.txt.UBound + 1)
                Set ctl = frmDesign.txt(frmDesign.txt.UBound)
            Else
                Set ctl = frmDesign.txt(Index)
            End If
    End Select
    '设置容器
    Set ctl.Container = frmDesign.picForm
    '设置位置
    ctl.Left = Left
    ctl.Top = Top
    If Width > 0 Then
        ctl.Width = Width
    End If
    If Height > 0 And ControlType <> "ComboBox" Then
        ctl.Height = Height
    End If
    '新增控件都与主窗口字体相同
    SetFont ctl, frmDesign.picForm

    Set LoadControl = ctl
End Function

Private Sub InitCont()
    Dim rsTmp As ADODB.Recordset, strSQL As String
    
    On Error GoTo errHandle
    

    '设置开单科室
    cmb开单.Clear
    cmb开单.AddItem "未指定"
    cmb开单.AddItem "病人所在科室"
    cmb开单.AddItem "操作员所在科室"
    cmb开单.AddItem "──────────────"
    
    Set rsTmp = GetDepartments("'临床','手术'", "1,2,3")
    Do Until rsTmp.EOF
        cmb开单.AddItem rsTmp("编码") & "-" & rsTmp("名称")
        cmb开单.ItemData(cmb开单.NewIndex) = rsTmp("ID")
        rsTmp.MoveNext
    Loop
    rsTmp.Close
    '设置可选项目列表
    lst项目.Clear
    lst项目.AddItem "病人ID"
    lst项目.AddItem "标识号"
    lst项目.AddItem "入院次数"
    lst项目.AddItem "性别"
    lst项目.AddItem "年龄"
    lst项目.AddItem "床号"
    lst项目.AddItem "病人病区"
    lst项目.AddItem "病人科室"
    lst项目.AddItem "费别"
    lst项目.AddItem "加班标志"
    lst项目.AddItem "婴儿费"
    lst项目.AddItem "应收合计"
    lst项目.AddItem "实收合计"
    '设置收费项目表格
    With msh收费项目
        .Rows = 13
        .ColWidth(0) = 1300
        .ColWidth(1) = 1000
        .ColAlignmentFixed(0) = 1
        .ColAlignment(1) = 1
        .Row = 0: .Col = 0: .Text = "栏目": .CellAlignment = 4
        .Row = 0: .Col = 1: .Text = "定义(1)": .CellAlignment = 4
        .Col = 1
        .TextMatrix(row收费类别值, 0) = "收费类别值"
        .TextMatrix(row收费细目值, 0) = "收费细目值"
        .TextMatrix(row数次值, 0) = "数次值"
        .TextMatrix(row收费类别, 0) = "收费类别"
        .TextMatrix(row细目选择, 0) = "细目选择"
        .TextMatrix(row计算单位, 0) = "计算单位"
        .TextMatrix(row数次, 0) = "数次"
        .TextMatrix(row标准单价, 0) = "标准单价"
        .TextMatrix(row应收金额, 0) = "应收金额"
        .TextMatrix(row实收金额, 0) = "实收金额"
        .TextMatrix(row执行部门, 0) = "执行部门"
        .TextMatrix(row附加标志, 0) = "附加标志"
    End With
    
    strSQL = "Select 编码,名称 as 类别 From 收费项目类别 Where 编码 Not In('1','4','5','6','7') Order by 序号"
    Set rsTmp = zlDatabase.OpensqlRecord(strSQL, Me.Caption)
    
    cmb类别.Clear
    cmb类别.AddItem "0-未指定"
    Do Until rsTmp.EOF
        cmb类别.AddItem rsTmp("编码") & "-" & rsTmp("类别")
        rsTmp.MoveNext
    Loop
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cmdSelect_Click()
    Dim rsTmp As New ADODB.Recordset
    Dim str类别 As String, lng项目ID As Long
    Dim strSQL As String
    
    str类别 = Left(msh收费项目.TextMatrix(row收费类别值, msh收费项目.Col), 1)
    If str类别 = "0" Then str类别 = ""
    If str类别 <> "" Then str类别 = "'" & str类别 & "'"
    
    lng项目ID = frmItemSelect.ShowSelect(Me, gstrPrivs, 0, 0, str类别)
    If lng项目ID <> 0 Then
        strSQL = "Select A.ID,A.类别||'-'||B.名称 as 类别,A.名称 From 收费项目目录 A,收费项目类别 B Where A.类别=B.编码 And A.ID=[1]"
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpensqlRecord(strSQL, Me.Caption, lng项目ID)
        msh收费项目.Text = rsTmp!名称
        msh收费项目.ColData(msh收费项目.Col) = rsTmp!ID
        msh收费项目.TextMatrix(row收费类别值, msh收费项目.Col) = rsTmp!类别
    End If
    msh收费项目.SetFocus
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    If Not InDesign Then
        glngOldProc = GetWindowLong(Me.hwnd, GWL_WNDPROC)
        Call SetWindowLong(Me.hwnd, GWL_WNDPROC, AddressOf wndProc)
    End If
End Sub

Private Sub Form_Resize()
    cmdHelp.Top = ScaleHeight - cmdHelp.Height - 200
    cmdOK.Top = cmdHelp.Top
    cmdCancel.Top = cmdHelp.Top
    
    cmdCancel.Left = ScaleWidth - cmdCancel.Width - 200
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - 100
    
    fra可选.Height = cmdHelp.Top - fra可选.Top - 100
    lst项目.Height = fra可选.Height - 300
    msh收费项目.Height = cmdOK.Top - msh收费项目.Top - 100
    msh收费项目.Width = ScaleWidth - msh收费项目.Left - 60
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange = True Then
        If MsgBox("如果你就这样退出的话，所有的修改都不会生效。" & vbCrLf & "是否确认退出？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1
            Exit Sub
        End If
    End If
    If Not InDesign Then
        Call SetWindowLong(Me.hwnd, GWL_WNDPROC, glngOldProc)
    End If
End Sub

Private Sub lst项目_ItemCheck(Item As Integer)
    mblnChange = True
End Sub

Private Sub msh收费项目_DblClick()
    Dim lngRow As Long, lngCol As Long
        
    With msh收费项目
        lngRow = .Row
        lngCol = .Col
        If lngCol < 1 Or lngCol > .Cols - 1 Then Exit Sub
        If lngRow < 2 Or lngRow > .Rows - 1 Then Exit Sub
        msh收费项目_KeyPress vbKeySpace
    End With
End Sub

Private Sub msh收费项目_GotFocus()
    If msh收费项目.Row = 1 Then
        cmb类别.Visible = True
        cmb类别.SetFocus
    End If
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Input细目() = True Then
            txtEdit.Visible = False
            KeyAscii = 0
        End If
    ElseIf KeyAscii = vbKeyEscape Then
        txtEdit.Visible = False
        msh收费项目.SetFocus
    End If
    
End Sub

Private Sub txtEdit_Validate(Cancel As Boolean)
    If txtEdit.Visible = False Then Exit Sub
    If txtEdit.Text = "" Then
        txtEdit.Visible = False
    Else
        If Input细目 = False Then
            Beep
            Cancel = True
        Else
            txtEdit.Visible = False
        End If
    End If
End Sub

Private Function Input细目() As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim str类别 As String, strSQL As String
    Dim str输入 As String, lng项目ID As Long
    
    str输入 = UCase(Replace(txtEdit.Text, "'", "''"))
    str类别 = Left(msh收费项目.TextMatrix(row收费类别值, msh收费项目.Col), 1)
    If str类别 = "0" Then str类别 = ""
    If str类别 <> "" Then str类别 = "'" & str类别 & "'"
    
    lng项目ID = frmItemSelect.ShowSelect(Me, gstrPrivs, 0, str类别, str输入, txtEdit.hwnd)
    If lng项目ID <> 0 Then
        strSQL = "Select ID,类别,名称 From 收费项目目录 Where ID=[1]"
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpensqlRecord(strSQL, Me.Caption, lng项目ID)
        With msh收费项目
            .ColData(.Col) = rsTmp!ID
            .TextMatrix(row收费细目值, .Col) = rsTmp!名称
            .TextMatrix(row收费类别值, .Col) = GetClassName(rsTmp!类别)
            .Row = row数次值
        End With
        mblnChange = True
        Input细目 = True
    Else
        zlControl.TxtSelAll txtEdit
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub txt数次_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Input数次 = True Then
            KeyAscii = 0
        End If
    ElseIf KeyAscii = vbKeyEscape Then
        txt数次.Visible = False
        msh收费项目.SetFocus
    ElseIf InStr("0123456789" & Chr(vbKeyBack), Chr(KeyAscii)) = 0 Then
        KeyAscii = 0
    End If
    
End Sub

Private Sub txt数次_Validate(Cancel As Boolean)
    If txt数次.Visible = False Then Exit Sub
    
    If txt数次.Text = "" Then
        txt数次.Visible = False
    Else
        If Input数次 = False Then
            Beep
            Cancel = True
        Else
            txt数次.Visible = False
        End If
    End If
End Sub

Private Function Input数次() As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strClass As String
    
    Dim strEdit As String
    
    
    strEdit = UCase(Replace(txt数次.Text, "'", "''"))
    If IsNumeric(strEdit) = False Then
        MsgBox "请输入合法的数值。", vbExclamation, gstrSysName
        txt数次.Text = ""
        Exit Function
    End If
    If Val(strEdit) < 0 Or Val(strEdit) > 1000 Then
        MsgBox "请输入1000以内的数。", vbExclamation, gstrSysName
        zlControl.TxtSelAll txt数次
        Exit Function
    End If
    With msh收费项目
        .TextMatrix(row数次值, .Col) = Format(strEdit, "0")
        If Val(strEdit) = 0 Then
            .TextMatrix(row数次, .Col) = "√"
        End If
        .Row = row收费类别
    End With
    txt数次.Visible = False
    mblnChange = True
    Input数次 = True
End Function

Private Sub txt数次_LostFocus()
    txt数次.Visible = False
End Sub

Private Sub txt项目数_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cmb开单_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub lst项目_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub msh收费项目_KeyPress(KeyAscii As Integer)
    With msh收费项目
        Select Case KeyAscii
            Case vbKeyReturn
                If .Row < .Rows - 1 Then
                    .Row = .Row + 1
                Else
                    If .Col = .Cols - 1 Then
                        SendKeys "{TAB}"
                    Else
                        .Row = 1: .Col = .Col + 1
                    End If
                End If
                Call msh收费项目_EnterCell
            Case vbKeySpace
                If .Row > row数次值 Then
                    If .Row = row收费类别 Then
                        If .TextMatrix(row收费类别值, .Col) = "" Or Mid(.TextMatrix(row收费类别值, .Col), 1, 1) = "0" Then
                            .TextMatrix(row收费类别, .Col) = "√"
                            Exit Sub
                        End If
                    ElseIf .Row = row数次 Then
                        If Val(.TextMatrix(row数次值, .Col)) = 0 Then
                            .TextMatrix(row数次, .Col) = "√"
                            Exit Sub
                        End If
                    End If
                    .TextMatrix(.Row, .Col) = IIf(.TextMatrix(.Row, .Col) = "√", "", "√")
                    mblnChange = True
                ElseIf .Row = row收费细目值 Then
                    txtEdit.Text = .Text
                    zlControl.TxtSelAll txtEdit
                    Call ShowTxtEdit
                ElseIf .Row = row数次值 Then
                    txt数次.Text = .Text
                    zlControl.TxtSelAll txt数次
                    Call Show数次
                End If
            Case Asc("*")
                If .Row = row收费细目值 Then
                    Call cmdSelect_Click
                End If
            Case Else
                If .Row = row收费细目值 Then
                    txtEdit.Text = Chr(KeyAscii)
                    txtEdit.SelStart = Len(txtEdit.Text)
                    Call ShowTxtEdit
                ElseIf .Row = row数次值 And InStr("0123456789", Chr(KeyAscii)) > 0 Then
                    txt数次.Text = Chr(KeyAscii)
                    txt数次.SelStart = Len(txt数次.Text)
                    Call Show数次
                End If
        End Select
    End With
End Sub

Private Sub msh收费项目_Scroll()
    cmdSelect.Visible = False
    cmb类别.Visible = False
    If msh收费项目.Row = 1 Then msh收费项目.Row = 2
'    Call msh收费项目_EnterCell
End Sub

Private Sub msh收费项目_EnterCell()
    Dim lngCount As Long
    
    cmb类别.Visible = False
    cmdSelect.Visible = False
    
    With msh收费项目
        If .Row = row收费类别值 Then
            cmb类别.Left = .Left + .CellLeft
            If .Row = row收费细目值 Then
                '发生了滚动，即调用了 msh收费项目_Scroll 事件
                cmdSelect.Left = .Left + .CellLeft + GetCellWidth - cmdSelect.Width
                cmdSelect.Visible = True
                Exit Sub
            End If
            cmb类别.Width = GetCellWidth()
            
            For lngCount = 0 To cmb类别.ListCount - 1
                If cmb类别.List(lngCount) = .Text Then
                    cmb类别.ListIndex = lngCount
                    Exit For
                End If
            Next
            If lngCount = cmb类别.ListCount Then
                cmb类别.ListIndex = 0
            End If
            cmb类别.Visible = True
            If cmb类别.Visible = True Then cmb类别.SetFocus
        ElseIf .Row = row收费细目值 Then
            cmdSelect.Left = .Left + .CellLeft + GetCellWidth - cmdSelect.Width
            cmdSelect.Visible = True
        End If
    End With
End Sub

Private Sub ShowTxtEdit()
    With msh收费项目
        cmdSelect.Visible = False
        txt数次.Visible = False
        txtEdit.Left = .Left + .CellLeft + 30
        txtEdit.Width = GetCellWidth() - 30
        txtEdit.Top = .Top + .CellTop + 45
        txtEdit.Visible = True
        txtEdit.SetFocus
    End With
End Sub

Private Sub Show数次()
    With msh收费项目
        cmdSelect.Visible = False
        txtEdit.Visible = False
        txt数次.Left = .Left + .CellLeft + 30
        txt数次.Width = GetCellWidth() - 30
        txt数次.Top = .Top + .CellTop + 45
        txt数次.ZOrder
        txt数次.Visible = True
        txt数次.SetFocus
    End With
End Sub

Private Function GetCellWidth() As Long
'得到当前单元格显示出来的宽度
    With msh收费项目
        If .CellLeft + .CellWidth > .Width Then
            '会出现纵向滚动条
            GetCellWidth = .Width - .CellLeft - 30
        Else
            GetCellWidth = .CellWidth - 30
        End If
    End With
    If GetCellWidth < 0 Then GetCellWidth = 0
End Function

Private Sub ud项目数_Change()
    Dim lngRow As Long
    
    With msh收费项目
        If .Cols < ud项目数.Value + 1 Then
            .Cols = ud项目数.Value + 1
            
            For lngRow = 1 To .Cols - 1
                .Row = 0: .Col = lngRow: .Text = "定义(" & lngRow & ")": .CellAlignment = 4
                .ColAlignment(.Col) = 1
                .ColWidth(.Col) = .ColWidth(.Col - 1)
            Next
            .Row = 1
            .TextMatrix(row收费类别值, ud项目数) = .TextMatrix(row收费类别值, ud项目数 - 1)
            For lngRow = row数次值 To .Rows - 1
                .TextMatrix(lngRow, ud项目数) = .TextMatrix(lngRow, ud项目数 - 1)
            Next
        Else
            .Cols = ud项目数.Value + 1
        End If
    End With
    mblnChange = True
    Call msh收费项目_EnterCell
End Sub

Private Function GetClassName(ByVal str编码 As String) As String
'根据类别编码，得到类别的全称
    Dim lngCount As Long
    For lngCount = 1 To cmb类别.ListCount - 1
        If Mid(cmb类别.List(lngCount), 1, 1) = str编码 Then
            GetClassName = cmb类别.List(lngCount)
            Exit Function
        End If
    Next
    GetClassName = cmb类别.List(0)
End Function

Private Function GetItemName(ByVal strID As String) As String
'根据类别编码，得到类别的全称
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo errH
    
    strSQL = "Select 名称 From 收费项目目录 Where ID=[1]"
    Set rsTmp = zlDatabase.OpensqlRecord(strSQL, Me.Caption, strID)
    If Not rsTmp.EOF Then GetItemName = rsTmp!名称
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
