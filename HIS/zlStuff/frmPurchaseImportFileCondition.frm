VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "VSFLEX8.OCX"
Begin VB.Form frmPurchaseImportFileCondition 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "条件设置"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8745
   Icon            =   "frmPurchaseImportFileCondition.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   8745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdExit 
      Caption         =   "退出"
      Height          =   300
      Left            =   7200
      TabIndex        =   5
      Top             =   240
      Width           =   885
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "保存"
      Height          =   300
      Left            =   6120
      TabIndex        =   4
      Top             =   240
      Width           =   885
   End
   Begin VB.OptionButton optFullImport 
      Caption         =   "完全导入"
      Height          =   255
      Left            =   3360
      TabIndex        =   2
      Top             =   240
      Width           =   1095
   End
   Begin VB.OptionButton optPartImport 
      Caption         =   "不完全导入"
      Height          =   255
      Left            =   1800
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfError 
      Height          =   4485
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   7215
      _cx             =   12726
      _cy             =   7911
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
      BackColorSel    =   12632256
      ForeColorSel    =   0
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
      Rows            =   17
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmPurchaseImportFileCondition.frx":000C
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
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
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   -1  'True
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
   Begin VB.Label lblImportMethod 
      AutoSize        =   -1  'True
      Caption         =   "导入方式"
      Height          =   180
      Left            =   360
      TabIndex        =   3
      Top             =   240
      Width           =   720
   End
End
Attribute VB_Name = "frmPurchaseImportFileCondition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const MCONFIXECOLOR As Long = &H8000000F  '不能修改列背景色
Private strPara As String   '参数值，规则为导入方式/卫材编码|数量|成本价|成本金额|发票金额|数量*成本价=成本金额|发票金额=成本金额|表格成本价=HIS成本价|效期|灭菌日期|灭菌效期|生产日期|存储库房|虚拟库房|商品条码(0-不完全导入1-完全导入/0-提示1-禁止|....)
Private mlngModal As Long '当前模块号

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim strTemp As String
    Dim intRow As Integer
    
    With vsfError
        If optFullImport.Value = True Then
            strTemp = "1/"
        Else
            strTemp = "0/"
        End If
        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 0) = "禁止" Then
                strTemp = strTemp & "1|"
            Else
                strTemp = strTemp & "0|"
            End If
        Next
    End With
    If strTemp <> "" Then
        strTemp = Mid(strTemp, 1, LenB(StrConv(strTemp, vbFromUnicode)) - 1)
    Else
        strTemp = "0/0|0|0|0|0|0|0|0|0|0|0|0|0|0|0"
    End If
    Call zlDatabase.SetPara("导入文件检查方式", strTemp, glngSys, mlngModal)
    Unload Me
End Sub

Private Sub Form_Load()
    Call InitControlPosition
    Call InitVSF
    Call LoadData
End Sub

Public Sub ShowMe(ByVal frmPar As Form, ByVal lngModal As Long)
    mlngModal = lngModal
    Me.Show vbModal, frmPar
End Sub

Private Sub InitControlPosition()
    '控件位置
    lblImportMethod.Move 70, 100
    optPartImport.Move lblImportMethod.Left + lblImportMethod.Width + 300, 100
    optFullImport.Move optPartImport.Left + optPartImport.Width + 150, 100
    cmdExit.Move Me.Width - cmdExit.Width - 100, lblImportMethod.Top - 50
    cmdSave.Move cmdExit.Left - cmdSave.Width - 100, lblImportMethod.Top - 50
    vsfError.Move lblImportMethod.Left, lblImportMethod.Top + lblImportMethod.Height + 150, Me.Width - 170, Me.Height - vsfError.Top - 150
End Sub

Private Sub InitVSF()
    '初始化vsf控件
    With vsfError
        .Editable = flexEDNone
        .SelectionMode = flexSelectionByRow
        .AllowUserResizing = flexResizeColumns
        .ExplorerBar = flexExNone
        .ExtendLastCol = True '最后一列填充满
        .ColComboList(0) = "禁止|提示"
        .WordWrap = True
        .AutoSize 2, 2, False, 0 = True
        .AutoSizeMode = flexAutoSizeRowHeight
        .ScrollBars = flexScrollBarVertical '将横向滚动条取消掉
        .Cell(flexcpBackColor, 1, 1, .Rows - 1, 2) = MCONFIXECOLOR '不能修改行颜色
    End With
End Sub

Private Sub optFullImport_Click()
    Dim intRow As Integer
    
    With vsfError
        If optFullImport.Value = True Then
            For intRow = 1 To .Rows - 1
                .TextMatrix(intRow, 0) = "禁止"
            Next
            .Cell(flexcpBackColor, 1, 0, .Rows - 1, 2) = MCONFIXECOLOR '不能修改行颜色
            .Editable = flexEDNone
        End If
    End With
End Sub

Private Sub optPartImport_Click()
    If optPartImport.Value = True Then
        vsfError.Cell(flexcpBackColor, 1, 0, vsfError.Rows - 1, 0) = &H80000005    '能修改行颜色
    End If
End Sub

Private Sub vsfError_CellChanged(ByVal Row As Long, ByVal Col As Long)
    With vsfError
        If Col = 0 Then
            If .TextMatrix(Row, Col) = "禁止" Then
                .Cell(flexcpFontBold, Row, 0, Row, 0) = True
            Else
                .Cell(flexcpFontBold, Row, 0, Row, 0) = False
            End If
        End If
    End With
End Sub

Private Sub vsfError_EnterCell()
    With vsfError
        If .Cell(flexcpBackColor, .Row, .Col, .Row, .Col) = MCONFIXECOLOR Then
            .Editable = flexEDNone
        Else
            .Editable = flexEDKbdMouse
        End If
    End With
End Sub

Private Sub LoadData()
    Dim strPara As String
    Dim intRow As Integer
    Dim intCol As Integer
    Dim arryPara As Variant
    Dim arryTempPara As Variant
    Dim strTemp As String
    Dim strImportMethod As String
    '加载数据
    If mlngModal = 1712 Then
        strPara = zlDatabase.GetPara("导入文件检查方式", glngSys, mlngModal, "0/0|0|0|0|0|0|0|0|0|0|0|0|0|0|0")
    Else
        strPara = zlDatabase.GetPara("导入文件检查方式", glngSys, mlngModal, "0/0|0|0|0|0|0|0|0|0|0|0|0|0|0|0")
    End If
    
    arryPara = Split(strPara, "|")
    With vsfError
        For intRow = 0 To UBound(arryPara)
            strTemp = arryPara(intRow)
            If intRow = 0 Then
                strImportMethod = Split(strTemp, "/")(0)
                If strImportMethod = "0" Then
                    optFullImport.Value = False
                    optPartImport.Value = True
                Else
                    optFullImport.Value = True
                    optPartImport.Value = False
                End If
                strTemp = Split(strTemp, "/")(1)
                strTemp = Split(strTemp, ",")(0)
                If strTemp = "0" Then
                    .TextMatrix(intRow + 1, 0) = "提示"
                Else
                    .TextMatrix(intRow + 1, 0) = "禁止"
                    .Cell(flexcpFontBold, intRow + 1, 0) = True
                End If
            End If
            If strTemp = "0" Then
                .TextMatrix(intRow + 1, 0) = "提示"
            Else
                .TextMatrix(intRow + 1, 0) = "禁止"
                .Cell(flexcpFontBold, intRow + 1, 0) = True
            End If
        Next
    End With
End Sub

