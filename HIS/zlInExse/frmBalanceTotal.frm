VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmBalanceTotal 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "结算信息汇总"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   5895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame 
      Height          =   6000
      Left            =   4500
      TabIndex        =   3
      Top             =   -360
      Width           =   10
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "退出(&C)"
      Height          =   350
      Left            =   4680
      TabIndex        =   2
      Top             =   150
      Width           =   1100
   End
   Begin VSFlex8Ctl.VSFlexGrid vsDeposit 
      Height          =   2100
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4245
      _cx             =   7488
      _cy             =   3704
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483634
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
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   350
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmBalanceTotal.frx":0000
      ScrollTrack     =   0   'False
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
   Begin VSFlex8Ctl.VSFlexGrid vsBalance 
      Height          =   2100
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3240
      Width           =   4245
      _cx             =   7488
      _cy             =   3704
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483634
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
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   350
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmBalanceTotal.frx":007A
      ScrollTrack     =   0   'False
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
   Begin XtremeSuiteControls.ShortcutCaption stcDeposit 
      Height          =   450
      Left            =   120
      TabIndex        =   5
      Top             =   40
      Width           =   4245
      _Version        =   589884
      _ExtentX        =   7488
      _ExtentY        =   794
      _StockProps     =   6
      Caption         =   "预交列表汇总"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
   Begin XtremeSuiteControls.ShortcutCaption stcBalance 
      Height          =   450
      Left            =   120
      TabIndex        =   4
      Top             =   2805
      Width           =   4245
      _Version        =   589884
      _ExtentX        =   7488
      _ExtentY        =   794
      _StockProps     =   6
      Caption         =   "支付列表汇总"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SubItemCaption  =   -1  'True
   End
End
Attribute VB_Name = "frmBalanceTotal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Private mstrDeposit As String '本次使用预交款的结算方式：结算方式1,结算方式2,结算方式3.....
Private mstrBalance As String '本次缴款使用的结算方式：结算方式1,结算方式2,结算方式3.....

Public Function ShowMe(ByVal frmParent As Object, ByVal vsDeposit As VSFlexGrid, ByVal vsBalance As VSFlexGrid) As Boolean
    '-------------------------------------------------------------------------------------------------------------------------
    '功能:显示预交列表汇总和结算列表汇总(结帐界面调用)
    '入参:frmParent-调用的父窗体
    '       vsDeposit-结帐界面的预交款列表
    '       vsBalance-结帐界面的结算列表
    '返回:调用成功返回True,否则返回False
    '-------------------------------------------------------------------------------------------------------------------------
    Call LoadDepositTotal(vsDeposit)
    Call LoadBalanceTotal(vsBalance)
    On Error Resume Next
    Me.Show 1, frmParent
End Function

Private Sub LoadDepositTotal(ByVal vsDeposit_In As VSFlexGrid)
    '--------------------------------------------------------------------------------------------------------------------------
    '功能:根据传入的vsDeposit_In列表,加载预交汇总列表
    '入参:vsDeposit_In-结帐界面的预交款列表
    '--------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, j As Integer
    Dim str结算方式 As String, strDeposit As String
    Dim dblMoney As Double, dblTotal As Double, strTmp As String
    Dim var结算方式 As Variant, varData As Variant, varTmp As Variant
    
    On Error GoTo errHandle
    
    With vsDeposit_In
        If .ColIndex("结算方式") = -1 Or .ColIndex("余额") = -1 Or .ColIndex("冲预交") = -1 Then Exit Sub
        For i = 1 To .Rows - 1
            If InStr("," & str结算方式 & ",", .TextMatrix(i, .ColIndex("结算方式"))) = 0 Then
                str结算方式 = str结算方式 & "," & .TextMatrix(i, .ColIndex("结算方式"))
            End If
            strDeposit = strDeposit & "|" & .TextMatrix(i, .ColIndex("结算方式")) & "," & Val(.TextMatrix(i, .ColIndex("余额"))) & "," & Val(.TextMatrix(i, .ColIndex("冲预交")))
        Next
    End With
    
    str结算方式 = Mid(str结算方式, 2): strDeposit = Mid(strDeposit, 2)
    If str结算方式 = "" Or strDeposit = "" Then Exit Sub
    var结算方式 = Split(str结算方式, ","): varData = Split(strDeposit, "|")
    
    For i = 0 To UBound(var结算方式)
        dblMoney = 0: dblTotal = 0
        For j = 0 To UBound(varData)
            varTmp = Split(varData(j), ",")
            If var结算方式(i) = varTmp(0) Then
                dblTotal = dblTotal + Val(varTmp(1))
                dblMoney = dblMoney + Val(varTmp(2))
            End If
        Next
        strTmp = strTmp & "|" & var结算方式(i) & "," & dblTotal & "," & dblMoney
    Next

    strTmp = Mid(strTmp, 2): If strTmp = "" Then Exit Sub
    varData = Split(strTmp, "|")
        
    With vsDeposit
        .Rows = UBound(varData) + 2
        For i = 1 To UBound(varData) + 1
            varTmp = Split(varData(i - 1), ",")
            .TextMatrix(i, .ColIndex("结算方式")) = varTmp(0)
            .TextMatrix(i, .ColIndex("余额")) = Format(Val(varTmp(1)), "0.00")
            .TextMatrix(i, .ColIndex("冲预交")) = Format(Val(varTmp(2)), "0.00")
        Next
        .ColWidth(.ColIndex("结算方式")) = IIf(.Rows <= 6, 1425, 1200)
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub LoadBalanceTotal(ByVal vsBalance_In As VSFlexGrid)
    '--------------------------------------------------------------------------------------------------------------------------
    '功能:根据传入的vsBalance_In列表,加载缴款汇总列表
    '入参:vsBalance_In-结帐界面的结算列表
    '--------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, j As Integer
    Dim str结算方式 As String, strBalance As String
    Dim dblMoney As Double, strTmp As String
    Dim var结算方式 As Variant, varData As Variant, varTmp As Variant
    On Error GoTo errHandle
    
    With vsBalance_In
        If .ColIndex("结算方式") = -1 Or .ColIndex("结算金额") = -1 Then Exit Sub
        For i = 1 To .Rows - 1
            If InStr("," & str结算方式 & ",", .TextMatrix(i, .ColIndex("结算方式"))) = 0 Then
                str结算方式 = str结算方式 & "," & .TextMatrix(i, .ColIndex("结算方式"))
            End If
            strBalance = strBalance & "|" & .TextMatrix(i, .ColIndex("结算方式")) & "," & Val(.TextMatrix(i, .ColIndex("结算金额")))
        Next
    End With
    
    str结算方式 = Mid(str结算方式, 2): strBalance = Mid(strBalance, 2)
    If str结算方式 = "" Or strBalance = "" Then Exit Sub
    var结算方式 = Split(str结算方式, ","): varData = Split(strBalance, "|")
    
    For i = 0 To UBound(var结算方式)
        dblMoney = 0
        For j = 0 To UBound(varData)
            varTmp = Split(varData(j), ",")
            If var结算方式(i) = varTmp(0) Then
                dblMoney = dblMoney + Val(varTmp(1))
            End If
        Next
        strTmp = strTmp & "|" & var结算方式(i) & "," & dblMoney
    Next

    strTmp = Mid(strTmp, 2): If strTmp = "" Then Exit Sub
    varData = Split(strTmp, "|")
        
    With vsBalance
        .Rows = UBound(varData) + 2
        For i = 1 To UBound(varData) + 1
            varTmp = Split(varData(i - 1), ",")
            .TextMatrix(i, .ColIndex("结算方式")) = varTmp(0)
            .TextMatrix(i, .ColIndex("支付金额")) = Format(Val(varTmp(1)), "0.00")
        Next
        .ColWidth(.ColIndex("结算方式")) = IIf(.Rows <= 6, 2310, 2085)
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
