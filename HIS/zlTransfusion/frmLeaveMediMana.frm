VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmLeaveMediMana 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "暂存药品管理"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9450
   Icon            =   "frmLeaveMediMana.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   9450
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmdCancle 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   8055
      TabIndex        =   2
      Top             =   5190
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   6765
      TabIndex        =   1
      Top             =   5190
      Width           =   1100
   End
   Begin VB.TextBox txtMain 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "gg yyyy""斥"" MM""岿"" dd""老"""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   3
      EndProperty
      Height          =   300
      Index           =   9
      Left            =   6630
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   4785
      Width           =   2535
   End
   Begin VB.TextBox txtMain 
      Height          =   300
      Index           =   8
      Left            =   3840
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   4785
      Width           =   1470
   End
   Begin VB.TextBox txtMain 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """￥""#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2052
         SubFormatType   =   2
      EndProperty
      Height          =   300
      Index           =   7
      Left            =   750
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   4785
      Width           =   1815
   End
   Begin VB.TextBox txtMain 
      Height          =   300
      Index           =   6
      Left            =   750
      MaxLength       =   200
      TabIndex        =   0
      Top             =   4440
      Width           =   8400
   End
   Begin VB.TextBox txtMain 
      Height          =   300
      Index           =   5
      Left            =   7365
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   495
      Width           =   1755
   End
   Begin VB.TextBox txtMain 
      Height          =   300
      Index           =   4
      Left            =   7365
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   825
      Width           =   1770
   End
   Begin VB.TextBox txtMain 
      Height          =   300
      Index           =   0
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   840
      Width           =   1065
   End
   Begin VB.TextBox txtMain 
      Height          =   300
      Index           =   1
      Left            =   2580
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   825
      Width           =   1245
   End
   Begin VB.TextBox txtMain 
      Height          =   300
      Index           =   2
      Left            =   4425
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   825
      Width           =   600
   End
   Begin VB.TextBox txtMain 
      Height          =   300
      Index           =   3
      Left            =   5655
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   825
      Width           =   540
   End
   Begin VSFlex8Ctl.VSFlexGrid vsList 
      Height          =   3180
      Left            =   45
      TabIndex        =   23
      Top             =   1185
      Width           =   9345
      _cx             =   16484
      _cy             =   5609
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
      BackColorSel    =   16772055
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
      RowHeightMax    =   2000
      ColWidthMin     =   0
      ColWidthMax     =   5000
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmLeaveMediMana.frx":6852
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   0   'False
      AutoSizeMode    =   1
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
      OwnerDraw       =   1
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
      Begin VB.TextBox txtEdit 
         Height          =   375
         Left            =   6120
         TabIndex        =   24
         Top             =   2490
         Visible         =   0   'False
         Width           =   1125
      End
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "[单位名称]暂存药品登记单"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   2595
      TabIndex        =   22
      Top             =   105
      Width           =   3795
   End
   Begin VB.Label lblMain 
      Caption         =   "登记时间"
      Height          =   240
      Index           =   9
      Left            =   5850
      TabIndex        =   21
      Top             =   4845
      Width           =   735
   End
   Begin VB.Label lblMain 
      Caption         =   "填制人"
      Height          =   225
      Index           =   8
      Left            =   3270
      TabIndex        =   19
      Top             =   4845
      Width           =   570
   End
   Begin VB.Label lblMain 
      Caption         =   "合计"
      Height          =   240
      Index           =   7
      Left            =   285
      TabIndex        =   17
      Top             =   4845
      Width           =   390
   End
   Begin VB.Label lblMain 
      Caption         =   "摘要"
      Height          =   240
      Index           =   6
      Left            =   270
      TabIndex        =   15
      Top             =   4500
      Width           =   390
   End
   Begin VB.Label lblMain 
      Caption         =   "NO"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Index           =   5
      Left            =   6975
      TabIndex        =   14
      Top             =   525
      Width           =   330
   End
   Begin VB.Label lblMain 
      Caption         =   "接收科室"
      Height          =   240
      Index           =   4
      Left            =   6540
      TabIndex        =   12
      Top             =   885
      Width           =   720
   End
   Begin VB.Label lblMain 
      Caption         =   "门诊号"
      Height          =   240
      Index           =   0
      Left            =   240
      TabIndex        =   10
      Top             =   885
      Width           =   570
   End
   Begin VB.Label lblMain 
      Caption         =   "姓名"
      Height          =   240
      Index           =   1
      Left            =   2100
      TabIndex        =   9
      Top             =   885
      Width           =   405
   End
   Begin VB.Label lblMain 
      Caption         =   "性别"
      Height          =   240
      Index           =   2
      Left            =   3990
      TabIndex        =   8
      Top             =   885
      Width           =   405
   End
   Begin VB.Label lblMain 
      Caption         =   "年龄"
      Height          =   240
      Index           =   3
      Left            =   5190
      TabIndex        =   7
      Top             =   885
      Width           =   405
   End
End
Attribute VB_Name = "frmLeaveMediMana"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum tMain
    门诊号 = 0
    姓名 = 1
    性别 = 2
    年龄 = 3
    暂存科室 = 4
    NO = 5
    摘要 = 6
    合计 = 7
    填制人 = 8
    填制日期 = 9
End Enum

Public pMediMaster As New MediMaster
Public pintType As Integer  '状态: 0-查看 1-增加 2-修改 3-消耗登记

Dim fntStrike As StdFont  '删除线

Private Sub init_vsList()
    Dim strHead As String
    If pintType = 1 Then
        '增加
        strHead = "药品来源,900,1;药品名称与编码,2500,1;规格,1600,1;用途,550,1;数量,750,7;计算单位,450,4;单价,750,7;金额,1000,7;" & _
                  "药品ID,0,1;医嘱ID,0,1;发送号,0,1;剂量单位,0,1;剂量系数,0,1;门诊单位,0,1;门诊包装,0,1;容量,0,1;可存数量,0,1"
    ElseIf pintType = 2 Then
        '修改
        strHead = "药品来源,900,1;药品名称与编码,2500,1;规格,1600,1;用途,550,1;数量,750,7;计算单位,450,4;单价,750,7;金额,1000,7;" & _
                  "药品ID,0,1;医嘱ID,0,1;发送号,0,1;剂量单位,0,1;剂量系数,0,1;门诊单位,0,1;门诊包装,0,1;容量,0,1;可存数量,0,1;UPDATE,0,1"
    ElseIf pintType = 3 Then
        '消耗登记
        strHead = "药品来源,900,1;药品名称与编码,2500,1;规格,1600,1;用途,550,1;可用数量,900,7;使用数量,900,7;计算单位,450,4;单价,0,7;金额,0,7;" & _
                  "药品ID,0,1;医嘱ID,0,1;发送号,0,1;剂量单位,0,1;剂量系数,0,1;门诊单位,0,1;门诊包装,0,1;容量,0,1;可存数量,0,1;数量,0,7;Key,0,1"
    Else
        '查看
    End If

    vsList.Redraw = flexRDNone
    Call SetVsFlexGridHead(strHead, vsList)
    With vsList
        .ColDataType(.ColIndex("数量")) = flexDTCurrency
        .ColFormat(.ColIndex("数量")) = "0.00"
        .ColDataType(.ColIndex("单价")) = flexDTCurrency
        .ColFormat(.ColIndex("单价")) = "0.00"
        .ColDataType(.ColIndex("金额")) = flexDTCurrency
        .ColFormat(.ColIndex("金额")) = "0.00"
        .Redraw = True
    End With


End Sub

Private Sub cmdCancle_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    Dim iRow As Integer, strErr As String, blnErr As Boolean, int序号 As Integer
    Dim blnDelRow As Boolean, iCol As Integer
    strErr = ""
    int序号 = 0
    With vsList
        '---删除空行
        
        .Select .Rows - 1, .Cols - 1
        blnDelRow = True
        Do While blnDelRow = True
            blnDelRow = False
            For iRow = 1 To .Rows - 1
                Select Case .TextMatrix(iRow, .ColIndex("药品来源"))
                Case "医嘱"
                    If Val(.TextMatrix(iRow, .ColIndex("医嘱ID"))) = 0 Or Val(.TextMatrix(iRow, .ColIndex("药品ID"))) = 0 Then
                        .RemoveItem iRow
                        blnDelRow = True
                        Exit For
                    End If
                Case "目录内"
                    If Val(.TextMatrix(iRow, .ColIndex("药品ID"))) = 0 Then
                        .RemoveItem iRow
                        blnDelRow = True
                        Exit For
                    End If
                Case "目录外"
                    If Val(.TextMatrix(iRow, .ColIndex("金额"))) = 0 Then
                        .RemoveItem iRow
                        blnDelRow = True
                        Exit For
                    End If
                Case Else
                    .RemoveItem iRow
                    blnDelRow = True
                    Exit For
                End Select
            Next
        Loop
        
        For iRow = 1 To .Rows - 1
            blnErr = False
            Select Case .TextMatrix(iRow, .ColIndex("药品来源"))
            Case "医嘱"
                If Val(.TextMatrix(iRow, .ColIndex("医嘱ID"))) = 0 Then
                    strErr = strErr & "第[" & iRow & "]数据有误，未指定医嘱。" & vbNewLine
                    blnErr = True
                End If
                If Val(.TextMatrix(iRow, .ColIndex("药品ID"))) = 0 Then
                    strErr = strErr & "第[" & iRow & "]数据有误，未指定药品。" & vbNewLine
                    blnErr = True
                End If
            Case "目录内"
                If Val(.TextMatrix(iRow, .ColIndex("药品ID"))) = 0 Then
                    strErr = strErr & "第[" & iRow & "]数据有误，未指定药品。" & vbNewLine
                    blnErr = True
                End If
            Case "目录外"
                '#
            Case Else
                strErr = strErr & "第[" & iRow & "]数据有误，未指定药品来源。" & vbNewLine
                blnErr = True
            End Select

            If Val(.TextMatrix(iRow, .ColIndex("金额"))) <= 0 Then
                If .TextMatrix(iRow, .ColIndex("药品来源")) = "医嘱" Then
                    strErr = strErr & "第[" & iRow & "]数据有误，未收费。" & vbNewLine
                Else
                    strErr = strErr & "第[" & iRow & "]数据有误，金额须大于零。" & vbNewLine
                End If
                blnErr = True
            End If
            
            For iCol = .FixedCols To .Cols - 1
               .TextMatrix(iRow, iCol) = DelInvalidChar(.TextMatrix(iRow, iCol), "'")
            Next
        Next
    End With
    If strErr <> "" Then
        MsgBox strErr, vbQuestion, gstrSysName
        Exit Sub
    End If

    If pintType = 1 Then
        Call AddLeveMedi
    ElseIf pintType = 2 Then
        Call UpdateLeveMedi
    ElseIf pintType = 3 Then
        Call UsedLeveMedi
    
    End If
    '退出
    Unload Me
End Sub

Private Sub UsedLeveMedi()
    Dim iRow As Integer, dbl改变量 As Double, curDate As Date
    If pintType <> 3 Then Exit Sub
    curDate = zlDatabase.Currentdate
    With vsList
        For iRow = 1 To .Rows - 1
            If .TextMatrix(iRow, .ColIndex("药品来源")) Like "目录*" Then
                dbl改变量 = Val(.TextMatrix(iRow, .ColIndex("使用数量")))
                If dbl改变量 > 0 Then
                    '直接增加
                    pMediMaster.摘要 = txtMain(6)
                    Call pMediMaster.InsertUseBill(.TextMatrix(iRow, .ColIndex("Key")), dbl改变量, curDate)
                End If
            End If
        Next
    End With
    
End Sub

Private Sub UpdateLeveMedi()
    Dim i As Integer
    If pintType = 2 Then
        '修改模式
        For i = 1 To pMediMaster.BillCount
            pMediMaster.Remove 1
        Next
        Call pMediMaster.DeleteBill(1)
        Call AddLeveMedi
    End If
End Sub

Private Sub AddLeveMedi()
    '检查数据正确性
    Dim iRow As Integer, strNO As String, int序号 As Integer, int用途 As Integer
    Dim objBIll As MediBill
    
    On Error GoTo errHandle
    With vsList
        For iRow = 1 To .Rows - 1
            Set objBIll = New MediBill
            int序号 = int序号 + 1
            objBIll.单价 = Val(.TextMatrix(iRow, .ColIndex("单价")))
            objBIll.规格 = .TextMatrix(iRow, .ColIndex("规格"))
            objBIll.剂量单位 = .TextMatrix(iRow, .ColIndex("剂量单位"))
            objBIll.剂量系数 = Val(.TextMatrix(iRow, .ColIndex("剂量系数")))
            objBIll.金额 = Val(.TextMatrix(iRow, .ColIndex("金额")))
            objBIll.计算单位 = .TextMatrix(iRow, .ColIndex("计算单位"))
            objBIll.门诊单位 = .TextMatrix(iRow, .ColIndex("门诊单位"))
            objBIll.门诊包装 = Val(.TextMatrix(iRow, .ColIndex("门诊包装")))
            objBIll.容量 = Val(.TextMatrix(iRow, .ColIndex("容量")))
            objBIll.入出系数 = 1
            objBIll.使用状态 = 0
            objBIll.数量 = Val(.TextMatrix(iRow, .ColIndex("数量")))
            objBIll.序号 = int序号
            objBIll.药品ID = Val(.TextMatrix(iRow, .ColIndex("药品ID")))
            objBIll.药品名称 = .TextMatrix(iRow, .ColIndex("药品名称与编码"))
            objBIll.医嘱ID = Val(.TextMatrix(iRow, .ColIndex("医嘱ID")))
            objBIll.发送号 = Val(.TextMatrix(iRow, .ColIndex("发送号")))

            Select Case .TextMatrix(iRow, .ColIndex("用途"))
            Case "输液"
                int用途 = 1
            Case "注射"
                int用途 = 2
            Case "皮试"
                int用途 = 3
            Case Else
                int用途 = 0
            End Select
            objBIll.执行分类 = int用途
            '加到类中
            Call pMediMaster.AddBill(objBIll, int序号)
        Next

        '写到数据库中
        If pintType = 1 Then
            '新增时，取NO号
            strNO = pMediMaster.GetNextNo
        ElseIf pintType = 2 Then
            '修改时，用原NO号
            strNO = pMediMaster.NO
        End If

        If strNO <> "" Then
            pMediMaster.摘要 = Trim(Replace(txtMain(6), "'", ""))
            Call pMediMaster.InsertBill(strNO, zlDatabase.Currentdate)
            pMediMaster.NO = strNO
            txtMain(tMain.NO) = strNO
        Else
            MsgBox "单据号错误,不能保存数据!", vbQuestion, Me.Caption
            Exit Sub
        End If
        '保存数据
        
        .AutoSize 0, .Cols - 1

    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Form_Activate()
    Dim str单位名称 As String
    On Error GoTo errHandle
    str单位名称 = GetSetting("ZLSOFT", "注册信息", "单位名称", "")
    
    If pMediMaster.病人ID = 0 Then Unload Me
    If pintType <> 0 Then
        '增加,修改
        vsList.Editable = flexEDKbdMouse
    End If
    txtEdit.Visible = False
    txtMain(tMain.NO) = pMediMaster.NO
    txtMain(tMain.填制日期) = Format(pMediMaster.登记时间, "yyyy-MM-dd hh:mm:ss")
    txtMain(tMain.填制人) = pMediMaster.操作员
    txtMain(tMain.门诊号) = pMediMaster.门诊号
    txtMain(tMain.年龄) = pMediMaster.年龄
    txtMain(tMain.姓名) = pMediMaster.姓名
    txtMain(tMain.性别) = pMediMaster.性别
    txtMain(tMain.暂存科室) = pMediMaster.科室名称
    txtMain(tMain.摘要) = pMediMaster.摘要
    If pintType <> 1 Then
        Call init_vsList
        Call Fill_vslist
    Else
        If pintType = 1 Then
            vsList.Select 1, 1
        End If
    End If
    If pintType = 3 Then
        lblTitle.Caption = str单位名称 & "暂存药品使用单"
    Else
        lblTitle.Caption = str单位名称 & "暂存药品登记单"
    End If
    If Me.ScaleWidth - lblTitle.Width < 0 Then
        lblTitle.Left = 10
    Else
        lblTitle.Left = (Me.ScaleWidth - lblTitle.Width) / 2
    End If
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub Fill_vslist()
    Dim objMediBill As MediBill
    Dim i As Integer, str用途 As String, str来源 As String
    With vsList
        
        For i = 1 To pMediMaster.BillCount
            Set objMediBill = pMediMaster.BillItem(i)
                If objMediBill.入出系数 = 1 Then
                    Select Case objMediBill.执行分类
                    Case 1
                        str用途 = "输液"
                    Case 2
                        str用途 = "注射"
                    Case 3
                        str用途 = "皮试"
                    Case Else
                        str用途 = "治疗"
                    End Select
                    If objMediBill.药品ID = 0 And objMediBill.医嘱ID = 0 Then
                        str来源 = "目录外"
                    ElseIf objMediBill.药品ID <> 0 And objMediBill.医嘱ID = 0 Then
                        str来源 = "目录内"
                    ElseIf objMediBill.药品ID <> 0 And objMediBill.医嘱ID <> 0 Then
                        str来源 = "医嘱"
                    Else
                        str来源 = "错误"
                    End If

                    .TextMatrix(.Rows - 1, .ColIndex("药品来源")) = str来源
                    .TextMatrix(.Rows - 1, .ColIndex("药品名称与编码")) = objMediBill.药品名称
                    .TextMatrix(.Rows - 1, .ColIndex("规格")) = objMediBill.规格
                    .TextMatrix(.Rows - 1, .ColIndex("用途")) = str用途
                    .TextMatrix(.Rows - 1, .ColIndex("数量")) = objMediBill.数量
                    .TextMatrix(.Rows - 1, .ColIndex("计算单位")) = objMediBill.计算单位
                    .TextMatrix(.Rows - 1, .ColIndex("单价")) = objMediBill.单价
                    .TextMatrix(.Rows - 1, .ColIndex("金额")) = objMediBill.金额
                    .TextMatrix(.Rows - 1, .ColIndex("药品ID")) = objMediBill.药品ID
                    .TextMatrix(.Rows - 1, .ColIndex("医嘱ID")) = objMediBill.医嘱ID
                    .TextMatrix(.Rows - 1, .ColIndex("发送号")) = objMediBill.发送号
                    .TextMatrix(.Rows - 1, .ColIndex("剂量单位")) = objMediBill.剂量单位
                    .TextMatrix(.Rows - 1, .ColIndex("剂量系数")) = objMediBill.剂量系数
                    .TextMatrix(.Rows - 1, .ColIndex("门诊单位")) = objMediBill.门诊单位
                    .TextMatrix(.Rows - 1, .ColIndex("门诊包装")) = objMediBill.门诊包装
                    .TextMatrix(.Rows - 1, .ColIndex("容量")) = objMediBill.容量
                    If pintType = 3 Then
                        .TextMatrix(.Rows - 1, .ColIndex("使用数量")) = 0
                        .TextMatrix(.Rows - 1, .ColIndex("可用数量")) = objMediBill.数量 - objMediBill.已用数量
                        .TextMatrix(.Rows - 1, .ColIndex("Key")) = objMediBill.序号 & "_" & objMediBill.入出系数 & "_" & Format(objMediBill.登记时间, "yyMMddhhmmss")
                        If objMediBill.数量 - objMediBill.已用数量 <= 0 Or str来源 = "医嘱" Then
                            .RemoveItem (.Rows - 1)
                        End If
                    End If
                    '.TextMatrix(.Rows - 1, .ColIndex("可存数量")) = objMediBill.可存数量
                    
                    .Rows = .Rows + 1
                End If
        Next
        If .Rows > 2 Then
            .RemoveItem (.Rows - 1)
        End If
        
        txtMain(tMain.合计) = Format(.Aggregate(flexSTSum, 1, .ColIndex("金额"), .Rows - 1, .ColIndex("金额")), "0.00")
        If pintType = 3 Then
            If .Rows < 3 Then
                If .TextMatrix(.Rows - 1, .ColIndex("药品名称与编码")) = "" Then
                    MsgBox "没有可登记的暂存药品!", vbInformation, Me.Caption
                    Me.cmdOk.Enabled = False
                End If
            End If
        End If
        .AutoSize 1, .Cols - 1
    End With
End Sub

Private Sub vsListButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, strInput As String
    Dim vRect As RECT, blnCanel As Boolean, strSelectRow As String '用于医嘱中排除已选择的记录
    Dim strPar As String, strType As String, i As Integer
    Dim strNO As String
    
    On Error GoTo errHandle
    
    If pintType = 0 Or pintType = 3 Then Exit Sub
    
    If Col = vsList.ColIndex("药品名称与编码") Then
        If vsList.TextMatrix(vsList.Row, vsList.ColIndex("药品来源")) = "目录外" Then Exit Sub

        '目录内
        '--------------------------------------------------------------------------------------
        If vsList.TextMatrix(vsList.Row, vsList.ColIndex("药品来源")) = "目录内" Then
            strInput = DelInvalidChar(UCase(Trim(txtEdit)), "'")
            If InStr(strInput, "]") > 0 Then
                strInput = Mid(Split(strInput, "]")(0), 2)
            End If
            If strInput = "" Then
                strSQL = "Select A.ID, A.编码, A.名称,A.规格, A.计算单位, B.现价, A.费用类型, Decode(A.服务对象, 1, '门诊', '门诊和住院') As 服务对象," & vbNewLine & _
                    "       A.执行科室,C.剂量系数,C.门诊单位,C.门诊包装,C.容量,D.剂量单位" & vbNewLine & _
                    "From 药品信息 D,药品规格 C,(Select 现价, 收费细目id,价格等级 From 收费价目 Where 终止日期 Is Null Or 终止日期 = To_Date('3000-01-01', 'YYYY-MM-DD')) B," & vbNewLine & _
                    "     收费项目目录 A" & vbNewLine & _
                    "Where C.药名ID=D.药名ID And A.Id=C.药品ID And A.ID = B.收费细目id And Mod(A.服务对象, 2) = 1 And" & vbNewLine & _
                    "      (A.撤档时间 Is Null Or A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) And" & vbNewLine & _
                    "      A.类别 In ( '5','6')" & GetPriceGradeSQL(gstr药品价格等级, gstr卫材价格等级, gstr普通项目价格等级, "A", "B", "1", "2", "3")
            Else
                strSQL = "Select A.ID, A.编码, A.名称,A.规格 , A.计算单位, B.现价, A.费用类型, Decode(A.服务对象, 1, '门诊', '门诊和住院') As 服务对象," & vbNewLine & _
                    "       A.执行科室,C.剂量系数,C.门诊单位,C.门诊包装,C.容量,D.剂量单位" & vbNewLine & _
                    "From 收费项目别名 E,药品信息 D,药品规格 C,(Select 现价, 收费细目id,价格等级 From 收费价目 Where 终止日期 Is Null Or 终止日期 = To_Date('3000-01-01', 'YYYY-MM-DD')) B," & vbNewLine & _
                    "     收费项目目录 A" & vbNewLine & _
                    "Where A.ID = E.收费细目id And  C.药名ID=D.药名ID And A.Id=C.药品ID And A.ID = B.收费细目id And Mod(A.服务对象, 2) = 1 And" & vbNewLine & _
                    "       (E.简码 Like '%" & strInput & "%' Or A.名称 Like '%" & strInput & "%' Or A.编码 Like '%" & strInput & "%') And E.码类 = 1 And" & vbNewLine & _
                    "      (A.撤档时间 Is Null Or A.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) And" & vbNewLine & _
                    "      A.类别 In ( '5','6')" & GetPriceGradeSQL(gstr药品价格等级, gstr卫材价格等级, gstr普通项目价格等级, "A", "B", "1", "2", "3")
            End If

            vRect = ZLControl.GetControlRect(txtEdit.hwnd)
            Set rsTmp = New ADODB.Recordset
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "药品", False, "", "选择药品", False, False, True, _
                                                 vRect.Left, vRect.Top, txtEdit.Height, blnCanel, True, True, gstr药品价格等级, gstr卫材价格等级, gstr普通项目价格等级)
            If Not blnCanel And rsTmp.State <> 0 Then
                If Not rsTmp.EOF Then
                    With vsList
                        .EditText = Replace("[" & zlCommFun.NVL(rsTmp.Fields("编码")) & "] " & zlCommFun.NVL(rsTmp.Fields("名称")), "[]", "")
                        .TextMatrix(.Row, .ColIndex("药品名称与编码")) = Replace("[" & zlCommFun.NVL(rsTmp.Fields("编码")) & "] " & zlCommFun.NVL(rsTmp.Fields("名称")), "[]", "")
                        .TextMatrix(.Row, .ColIndex("单价")) = Format(zlCommFun.NVL(rsTmp.Fields("现价"), 0), "0.00")
                        .TextMatrix(.Row, .ColIndex("计算单位")) = zlCommFun.NVL(rsTmp.Fields("计算单位"), "")
                        .TextMatrix(.Row, .ColIndex("规格")) = zlCommFun.NVL(rsTmp.Fields("规格"), "")
                        .TextMatrix(.Row, .ColIndex("药品ID")) = zlCommFun.NVL(rsTmp.Fields("ID"), "")
                        .TextMatrix(.Row, .ColIndex("医嘱ID")) = 0
                        .TextMatrix(.Row, .ColIndex("剂量单位")) = zlCommFun.NVL(rsTmp.Fields("剂量单位"), "")
                        .TextMatrix(.Row, .ColIndex("剂量系数")) = zlCommFun.NVL(rsTmp.Fields("剂量系数"), "")
                        .TextMatrix(.Row, .ColIndex("门诊单位")) = zlCommFun.NVL(rsTmp.Fields("门诊单位"), "")
                        .TextMatrix(.Row, .ColIndex("门诊包装")) = zlCommFun.NVL(rsTmp.Fields("门诊包装"), "")
                    End With
                End If
                Set rsTmp = Nothing
            End If
            txtEdit = ""
        End If

        '医嘱
        '--------------------------------------------------------------------------------------
        If vsList.TextMatrix(vsList.Row, vsList.ColIndex("药品来源")) = "医嘱" Then '
            strInput = DelInvalidChar(UCase(Trim(txtEdit)), "'")
            If InStr(strInput, "]") > 0 Then
                strInput = Mid(Split(strInput, "]")(0), 2)
            End If

            '取得已选择的行
            strSelectRow = ""
            For i = 1 To vsList.Rows - 1
                With vsList
                    If Val(.TextMatrix(i, .ColIndex("医嘱ID"))) > 0 Then
                        strSelectRow = strSelectRow & Val(.TextMatrix(i, .ColIndex("医嘱ID"))) & "_" & Val(.TextMatrix(i, .ColIndex("发送号"))) & ","
                    End If
                End With
            Next
            
            strPar = zlDatabase.GetPara("显示单据种类", glngSys, 1264, "1,1,1,1")
            For i = 0 To 3
                strType = strType & IIf(Val(Split(strPar, ",")(i)) = 1, "," & i, "")
            Next
            
            If pMediMaster.挂号单 Like "*_*" Then
                '门诊留观
                strNO = " a.主页id = " & Split(pMediMaster.挂号单, "_")(1) & " "
            Else
                '门诊
                strNO = " a.挂号单 = '" & pMediMaster.挂号单 & "' "
            End If
            
            If strInput = "" Then
                strSQL = "Select c.相关id, i.发送时间, i.No, b.执行分类 as 用途, i.发送号, i.医嘱id, d.编码, d.名称, d.规格, nvl(g.剂量单位,'') as 剂量单位, nvl(e.剂量系数,0) 剂量系数," & vbNewLine & _
                        "            nvl(e.门诊单位,'') as 门诊单位 , nvl(e.门诊包装,0) 门诊包装, nvl(e.容量,0) as 容量, h.标准单价 As 现价, (Nvl(i.发送数次, 0) / Nvl(e.剂量系数, 0)) As 可存数量, c.收费细目id," & vbNewLine & _
                        "            (Nvl(f.数量, 0) * Nvl(f.入出系数, 0)) As 已存数量, d.计算单位, C.单次用量, e.门诊可否分零 " & vbNewLine & _
                        "From 门诊费用记录 h," & vbNewLine & _
                        "        药品信息 g, 暂存药品记录 f, 药品规格 e, 收费项目目录 d, 病人医嘱记录 c, 诊疗项目目录 b, 病人医嘱发送 i," & vbNewLine & _
                        "        病人医嘱记录 a" & vbNewLine & _
                        "Where Instr('" & strType & "', nvl(b.执行分类,0))>0 And c.id = h.医嘱序号(+) And h.记录状态(+)=1 and h.费用状态(+)<>1 And e.药名id = g.药名id And Mod(d.服务对象, 2) = 1 And" & vbNewLine & _
                        "           (d.撤档时间 Is Null Or d.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) And " & vbNewLine & _
                        "           i.医嘱id = f.医嘱id(+) And i.发送号 = f.发送号(+) And c.收费细目id = e.药品id And c.收费细目id = d.Id And" & vbNewLine & _
                        "           (Nvl(I.发送数次, 0) / Nvl(E.剂量系数, 0)) - (Nvl(F.数量, 0) * Nvl(F.入出系数, 0)) > 0 And " & vbNewLine & _
                        "           c.诊疗类别 In ('5', '6') And a.Id = c.相关id And a.诊疗项目id = b.Id And i.医嘱id = c.Id And f.科室id(+) = " & pMediMaster.科室ID & " And" & vbNewLine & _
                        "           a.病人id = " & pMediMaster.病人ID & " And a.执行科室id = " & pMediMaster.科室ID & " And F.入出系数(+)=1 And a.诊疗类别 = 'E' And " & strNO & vbNewLine & _
                        IIf(strSelectRow = "", "", " And Instr('" & strSelectRow & "' ,I.医嘱ID||'_'||I.发送号||',')<=0 ") & vbNewLine & _
                        "Order By 发送号, 医嘱id,相关id"
            Else
                strSQL = "Select c.相关id, i.发送时间, i.No, b.执行分类 as 用途, i.发送号, i.医嘱id, d.编码, d.名称, d.规格, nvl(g.剂量单位,'') as 剂量单位, nvl(e.剂量系数,0) as 剂量系数," & vbNewLine & _
                        "            nvl(e.门诊单位,'') as 门诊单位, nvl(e.门诊包装,0) as 门诊包装, nvl(e.容量,0) as 容量, h.标准单价 As 现价, (Nvl(i.发送数次, 0) / Nvl(e.剂量系数, 0)) As 可存数量, c.收费细目id," & vbNewLine & _
                        "            (Nvl(f.数量, 0) * Nvl(f.入出系数, 0)) As 已存数量, d.计算单位,C.单次用量, e.门诊可否分零 " & vbNewLine & _
                        "From 门诊费用记录 h," & vbNewLine & _
                        "        药品信息 g, 暂存药品记录 f, 药品规格 e, 收费项目目录 d, 病人医嘱记录 c, 诊疗项目目录 b, 病人医嘱发送 i," & vbNewLine & _
                        "        病人医嘱记录 a" & vbNewLine & _
                        "Where  Instr('" & strType & "', nvl(b.执行分类,0))>0 And c.id = h.医嘱序号(+) And h.记录状态(+)=1 And h.费用状态(+)<>1 And e.药名id = g.药名id And Mod(d.服务对象, 2) = 1 And" & vbNewLine & _
                        "           (d.撤档时间 Is Null Or d.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) And" & vbNewLine & _
                        "                        (Zlspellcode(d.名称) Like '%" & strInput & "%' Or d.名称 Like '%" & strInput & "%' Or" & vbNewLine & _
                        "                        d.编码 Like '%" & strInput & "%') And" & vbNewLine & _
                        "           i.医嘱id = f.医嘱id(+) And i.发送号 = f.发送号(+) And c.收费细目id = e.药品id And c.收费细目id = d.Id And" & vbNewLine & _
                        "           (Nvl(I.发送数次, 0) / Nvl(E.剂量系数, 0)) - (Nvl(F.数量, 0) * Nvl(F.入出系数, 0)) > 0 And " & vbNewLine & _
                        "           c.诊疗类别 In ('5', '6') And a.Id = c.相关id And a.诊疗项目id = b.Id And i.医嘱id = c.Id And f.科室id(+) = " & pMediMaster.科室ID & " And" & vbNewLine & _
                        "           a.病人id = " & pMediMaster.病人ID & " And a.执行科室id = " & pMediMaster.科室ID & " And F.入出系数(+)=1 And a.诊疗类别 = 'E' And " & strNO & vbNewLine & _
                        IIf(strSelectRow = "", "", " And Instr('" & strSelectRow & "' ,I.医嘱ID||'_'||I.发送号||',')<=0 ") & vbNewLine & _
                        "Order By 发送号, 医嘱id,相关id"
            End If
            Call frmLeaveSelect.LeaveSelect(Me, strSQL)
            txtEdit = ""
       

        End If
    End If
    Call zlCommFun.PressKey(vbKeyRight)
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If

End Sub

Private Sub Form_Load()
    Call init_vsList
End Sub

Private Sub Form_Unload(Cancel As Integer)
    pintType = 0
    Set pMediMaster = Nothing
End Sub


Private Sub vsList_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim str来源
    If Col = vsList.ColIndex("药品来源") Then
        str来源 = vsList.TextMatrix(Row, Col)
        vsList.Delete
        vsList.TextMatrix(Row, Col) = str来源
    End If
End Sub

Private Sub vsList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)

    Dim blnEdit As Boolean
    Call vsList_BeforeEdit(NewRow, NewCol, blnEdit)
    If blnEdit Then
        vsList.ComboList = ""
        'vsList.FocusRect = flexFocusLight
    Else
        'vsList.FocusRect = flexFocusSolid
        If NewCol = vsList.ColIndex("药品名称与编码") Then
            vsList.ComboList = "..."
        ElseIf NewCol = vsList.ColIndex("药品来源") Then
            vsList.ComboList = "医嘱|目录内|目录外"
        ElseIf NewCol = vsList.ColIndex("用途") Then
            vsList.ComboList = "治疗|输液|注射|皮试"
        Else
            vsList.ComboList = ""
        End If
    End If

End Sub

Private Sub vsList_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strEditRow As String
    On Error GoTo errHandle

    If pintType = 1 Then
        With vsList
        Select Case .TextMatrix(Row, .ColIndex("药品来源"))
        Case "目录内"
            strEditRow = "," & .ColIndex("药品来源") & "," & .ColIndex("药品名称与编码") & "," & .ColIndex("数量") & "," & .ColIndex("用途") & ","
        Case "目录外"
            strEditRow = "," & .ColIndex("药品来源") & "," & .ColIndex("药品名称与编码") & "," & .ColIndex("规格") & "," & .ColIndex("数量") & "," & .ColIndex("单价") & "," & .ColIndex("计算单位") & "," & .ColIndex("用途") & ","
        Case Else
            strEditRow = "," & .ColIndex("药品来源") & "," & .ColIndex("药品名称与编码") & "," & .ColIndex("数量") & ","
        End Select
        End With
    ElseIf pintType = 2 Then
        With vsList
        Select Case .TextMatrix(Row, .ColIndex("药品来源"))
        Case "目录内"
            strEditRow = "," & .ColIndex("药品来源") & "," & .ColIndex("药品名称与编码") & "," & .ColIndex("数量") & "," & .ColIndex("用途") & ","
        Case "目录外"
            strEditRow = "," & .ColIndex("药品来源") & "," & .ColIndex("药品名称与编码") & "," & .ColIndex("规格") & "," & .ColIndex("数量") & "," & .ColIndex("单价") & "," & .ColIndex("计算单位") & "," & .ColIndex("用途") & ","
        Case Else
            strEditRow = "," & .ColIndex("药品来源") & "," & .ColIndex("药品名称与编码") & "," & .ColIndex("数量") & ","
        End Select
        End With
    ElseIf pintType = 3 Then
        strEditRow = "," & vsList.ColIndex("使用数量") & ","
    Else
        strEditRow = ""
    End If

    If InStr(strEditRow, "," & Col & ",") <= 0 Then
        Cancel = True
    End If
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If

End Sub


Private Sub vsList_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Call vsListButtonClick(Row, Col)
End Sub

Private Sub vsList_EnterCell()
    On Error GoTo errHandle
    With vsList
    
        If pintType = 0 Or pintType = 3 Then Exit Sub
        If .Col = .ColIndex("药品名称与编码") And .Row > 0 Then
            If txtEdit.Tag = "False" And InStr("目录内,医嘱", .TextMatrix(.Row, .ColIndex("药品来源"))) > 0 Then
                txtEdit.Left = .CellLeft
                txtEdit.Top = .CellTop
                txtEdit.Height = .CellHeight - 12
                txtEdit.Width = .CellWidth - 12
                txtEdit.Tag = "True"
            End If
        Else
            txtEdit.Tag = "False"
        End If
        Dim blnCancle As Boolean
        Call vsList_BeforeEdit(.Row, .Col, blnCancle)
        If Not blnCancle Then
            Call .CellBorder(vsList.GridColor, 1, 1, 2, 2, 0, 0)
        End If
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub vsList_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo errHandle
    Dim strLast药品来源 As String
    With vsList
        If KeyCode = vbKeyReturn Then
            If .Col = .ColIndex("金额") And .Row = .Rows - 1 And (pintType = 1 Or pintType = 2) Then
                strLast药品来源 = .TextMatrix(.Row, .ColIndex("药品来源"))

                .Rows = .Rows + 1
                .Row = .Row + 1

                If strLast药品来源 <> "" Then
                    .TextMatrix(.Row, .ColIndex("药品来源")) = strLast药品来源
                    .Col = .ColIndex("药品名称与编码")
                Else
                    .Col = .ColIndex("药品来源")
                End If
            Else
                If .Cols > .Col + 1 And .Col <> .ColIndex("金额") Then
                    .Col = .Col + 1
                Else
                    If .Rows > .Row + 1 Then
                        .Row = .Row + 1
                        .Col = .ColIndex("药品名称与编码")
                    End If
                End If
            End If
        End If
        If pintType = 1 Or pintType = 2 Then
            '增加
            If KeyCode = vbKeyDelete Then
                If MsgBox("是否删除当前行?", vbQuestion + vbYesNo + vbDefaultButton2, Me.Caption) = vbYes Then
                    If .Rows > 2 Then
                        .RemoveItem (.Row)
                    Else
                        .Delete
                    End If
                End If
            End If
            
        End If
    End With
    
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub vsList_KeyPress(KeyAscii As Integer)
    If pintType = 0 Or pintType = 3 Then Exit Sub
    With vsList
        If (.Col = .ColIndex("药品名称与编码")) And KeyAscii = vbKeyReturn Then
            KeyAscii = 0
        Else
            If .Col = .ColIndex("药品名称与编码") And vsList.ComboList = "..." Then
                If KeyAscii = Asc("*") Then
                    KeyAscii = 0
                    txtEdit.Text = .EditText
                    Call vsList_CellButtonClick(.Row, .Col)
                    txtEdit.Tag = False
                    txtEdit.Visible = False
                Else
                    .ComboList = "" '使按钮状态进入输入状态
                End If
            End If
        End If
    End With
End Sub

Private Sub vsList_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)

On Error GoTo errHandle
    If pintType = 0 Or pintType = 3 Then Exit Sub
    With vsList
        If Col = .ColIndex("药品名称与编码") And KeyAscii = vbKeyReturn And .TextMatrix(.Row, .ColIndex("药品来源")) <> "目录外" Then
            txtEdit.Text = .EditText
            .EditText = ""
            Call vsListButtonClick(Row, Col)
            txtEdit.Tag = False
            txtEdit.Visible = False
        ElseIf KeyAscii = vbKeyReturn Then
            If .Cols < .Col + 1 And .Col <> .ColIndex("金额") Then
                .Col = .Col + 1
            Else
                If .Rows < .Row + 1 Then
                    .Row = .Row + 1
                    .Col = .ColIndex("药品名称与编码")
                End If
            End If
        End If
    End With

    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub vsList_LeaveCell()
    With vsList
        On Error GoTo errHandle
        If pintType = 1 Or pintType = 2 Then
            If Val(.TextMatrix(.Row, .ColIndex("数量"))) <> 0 Then
                .TextMatrix(.Row, .ColIndex("金额")) = Format(Val(.TextMatrix(.Row, .ColIndex("数量"))) * Val(.TextMatrix(.Row, .ColIndex("单价"))), "0.00")
            Else
                .TextMatrix(.Row, .ColIndex("金额")) = "0.00"
            End If
        End If
        txtMain(tMain.合计) = Format(.Aggregate(flexSTSum, 1, .ColIndex("金额"), .Rows - 1, .ColIndex("金额")), "0.00")
        If .TextMatrix(.Row, .ColIndex("药品来源")) = "" Then .TextMatrix(.Row, .ColIndex("药品来源")) = "医嘱"
    
        Dim blnCancle As Boolean
        Call vsList_BeforeEdit(.Row, .Col, blnCancle)
        If Not blnCancle Then
            On Error Resume Next
            Call .CellBorder(vsList.GridColor, 0, 0, 0, 0, 0, 0)
        End If
        
    End With
    
    Exit Sub
errHandle:
    If Err.Number = 381 Then Exit Sub
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub vsList_RowColChange()

On Error GoTo errHandle
    With vsList
        If txtEdit.Tag = "True" Then
            txtEdit.Left = .CellLeft
            txtEdit.Top = .CellTop
            txtEdit.Height = .CellHeight - 12
            txtEdit.Width = .CellWidth - 12
        End If
        
        If pintType = 1 Or pintType = 2 Then
            If Val(.TextMatrix(.Row, .ColIndex("数量"))) <> 0 Then
                .TextMatrix(.Row, .ColIndex("金额")) = Format(Val(.TextMatrix(.Row, .ColIndex("数量"))) * Val(.TextMatrix(.Row, .ColIndex("单价"))), "0.00")
            Else
                .TextMatrix(.Row, .ColIndex("金额")) = "0.00"
            End If
        End If
        txtMain(tMain.合计) = Format(.Aggregate(flexSTSum, 1, .ColIndex("金额"), .Rows - 1, .ColIndex("金额")), "0.00")
    End With
    Exit Sub
errHandle:
    If Err.Number = 381 Then Exit Sub
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub vsList_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rsTmp As New ADODB.Recordset, strSQL As String
    Dim lng可存数量 As Long
    On Error GoTo errHandle
    With vsList
    Select Case Col
        Case .ColIndex("数量"), .ColIndex("单价"), .ColIndex("金额 ")
            If IsNumeric(.EditText) = False Then Cancel = True
            If .ColIndex("数量") Then
                '增加时,输入的数量不能大于可存数量
                If pintType = 1 And .TextMatrix(Row, .ColIndex("药品来源")) = "医嘱" Then
                    If Val(.EditText) > Val(.TextMatrix(Row, .ColIndex("可存数量"))) Then
                        MsgBox "数量填写错误，此药品最多只能寄存 " & Val(.TextMatrix(Row, .ColIndex("可存数量"))) & " " & .TextMatrix(Row, .ColIndex("计算单位")), vbQuestion, Me.Caption
                        Cancel = True
                    End If
                End If
                '修改时,重新取可存数量
                If pintType = 2 And .TextMatrix(Row, .ColIndex("药品来源")) = "医嘱" Then
'                    strSQL = "Select c.相关id, i.发送时间, i.No, b.执行分类 as 用途, i.发送号, i.医嘱id, d.编码, d.名称, d.规格, g.剂量单位, e.剂量系数," & vbNewLine & _
'                            "            e.门诊单位, e.门诊包装, e.容量, h.现价, (Nvl(i.发送数次, 0) / Nvl(e.剂量系数, 0)) As 可存数量, c.收费细目id," & vbNewLine & _
'                            "            (Nvl(f.数量, 0) * Nvl(f.入出系数, 0)) As 已存数量, d.计算单位" & vbNewLine & _
'                            "From (Select 现价, 收费细目id From 收费价目 Where 终止日期 Is Null Or 终止日期 = To_Date('3000-01-01', 'YYYY-MM-DD')) h," & vbNewLine & _
'                            "        药品信息 g, 暂存药品记录 f, 药品规格 e, 收费项目目录 d, 病人医嘱记录 c, 诊疗项目目录 b, 病人医嘱发送 i," & vbNewLine & _
'                            "        病人医嘱记录 a" & vbNewLine & _
'                            "Where c.收费细目id = h.收费细目id(+) And e.药名id = g.药名id And Mod(d.服务对象, 2) = 1 And" & vbNewLine & _
'                            "           (d.撤档时间 Is Null Or d.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) And Nvl(d.是否变价, 0) = 0 And" & vbNewLine & _
'                            "           i.医嘱id = f.医嘱id(+) And i.发送号 = f.发送号(+) And c.收费细目id = e.药品id And c.收费细目id = d.Id And" & vbNewLine & _
'                            "           c.诊疗类别 In ('5', '6') And a.Id = c.相关id And a.诊疗项目id = b.Id And i.医嘱id = c.Id And f.科室id(+) = [2] And" & vbNewLine & _
'                            "           a.病人id = [1] And a.执行科室id = [2] And a.诊疗类别 = 'E' And a.挂号单 = [3]" & vbNewLine & _
'                            " And F.入出系数(+)=1 And i.医嘱ID=[4] And i.发送号=[5]  " & vbNewLine & _
'                            "Order By 发送号, 医嘱id,相关id"
                    strSQL = "Select c.相关id, i.发送时间, i.No, b.执行分类 as 用途, i.发送号, i.医嘱id, d.编码, d.名称, d.规格, g.剂量单位, e.剂量系数," & vbNewLine & _
                            "            e.门诊单位, e.门诊包装, e.容量, h.标准单价 As 现价, (Nvl(i.发送数次, 0) / Nvl(e.剂量系数, 0)) As 可存数量, c.收费细目id," & vbNewLine & _
                            "            (Nvl(f.数量, 0) * Nvl(f.入出系数, 0)) As 已存数量, d.计算单位" & vbNewLine & _
                            "From 门诊费用记录 h," & vbNewLine & _
                            "        药品信息 g, 暂存药品记录 f, 药品规格 e, 收费项目目录 d, 病人医嘱记录 c, 诊疗项目目录 b, 病人医嘱发送 i," & vbNewLine & _
                            "        病人医嘱记录 a" & vbNewLine & _
                            "Where c.id = h.医嘱序号(+) And e.药名id = g.药名id And Mod(d.服务对象, 2) = 1 And" & vbNewLine & _
                            "           (d.撤档时间 Is Null Or d.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) And" & vbNewLine & _
                            "           i.医嘱id = f.医嘱id(+) And i.发送号 = f.发送号(+) And c.收费细目id = e.药品id And c.收费细目id = d.Id And" & vbNewLine & _
                            "           c.诊疗类别 In ('5', '6') And a.Id = c.相关id And a.诊疗项目id = b.Id And i.医嘱id = c.Id And f.科室id(+) = [2] And" & vbNewLine & _
                            "           a.病人id = [1] And a.执行科室id = [2] And a.诊疗类别 = 'E' And a.挂号单 = [3]" & vbNewLine & _
                            " And F.入出系数(+)=1 And i.医嘱ID=[4] And i.发送号=[5]  " & vbNewLine & _
                            "Order By 发送号, 医嘱id,相关id"
                     Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "重新取可存数量", pMediMaster.病人ID, pMediMaster.科室ID, pMediMaster.挂号单, Val(.TextMatrix(Row, .ColIndex("医嘱ID"))), Val(.TextMatrix(Row, .ColIndex("发送号"))))
                     If Not rsTmp.EOF Then
                        lng可存数量 = rsTmp.Fields("可存数量") - (rsTmp.Fields("已存数量") - Val(.TextMatrix(Row, .ColIndex("数量"))))
                        If Val(.EditText) > lng可存数量 Then
                            MsgBox "数量填写错误，此药品最多只能寄存 " & lng可存数量 & " " & .TextMatrix(Row, .ColIndex("计算单位")), vbQuestion, Me.Caption
                            Cancel = True
                        End If
                     End If
                End If

            End If
        Case .ColIndex("使用数量")
            If IsNumeric(.EditText) = True Then
                If Val(.TextMatrix(Row, .ColIndex("可用数量"))) < Val(.EditText) Or Val(.EditText) < 0 Then
                    If Val(.EditText) > 0 Then
                        MsgBox "使用数量不能大于可用数量!", vbQuestion, Me.Caption
                    ElseIf Val(.EditText) < 0 Then
                        MsgBox "使用数量不能小于0!", vbQuestion, Me.Caption
                    End If
                    Cancel = True
                End If
            Else
                Cancel = True
            End If
        
        Case .ColIndex("药品来源")
            If InStr(",医嘱,目录内,目录外,", "," & .EditText & ",") <= 0 Then Cancel = True
        Case .ColIndex("用途")
            If InStr(",治疗,输液,注射,皮试,", "," & .EditText & ",") <= 0 Then Cancel = True

    End Select
    End With
    Exit Sub
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
