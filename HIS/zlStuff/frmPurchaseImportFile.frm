VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPurchaseImportFile 
   Caption         =   "导入外部文件"
   ClientHeight    =   6630
   ClientLeft      =   120
   ClientTop       =   510
   ClientWidth     =   9075
   Icon            =   "frmPurchaseImportFile.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6630
   ScaleWidth      =   9075
   StartUpPosition =   1  '所有者中心
   Begin MSComctlLib.ImageList imgError 
      Left            =   5400
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseImportFile.frx":6852
            Key             =   "error"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchaseImportFile.frx":D0B4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picSplit 
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   240
      MousePointer    =   7  'Size N S
      ScaleHeight     =   300
      ScaleWidth      =   4215
      TabIndex        =   10
      Top             =   4560
      Width           =   4215
      Begin VB.Label lblCollect 
         AutoSize        =   -1  'True
         Caption         =   "成本金额：     元          发票金额：     元"
         Height          =   180
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   3960
      End
   End
   Begin MSComDlg.CommonDialog dlgOpenFile 
      Left            =   3360
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "*.xls|*.xls|*.xlsx|*.xlsx"
   End
   Begin MSComctlLib.ProgressBar ProCheck 
      Height          =   300
      Left            =   360
      TabIndex        =   7
      Top             =   6000
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.CommandButton cmdProvider 
      Caption         =   "…"
      Height          =   300
      Left            =   4440
      TabIndex        =   5
      Top             =   960
      Width           =   280
   End
   Begin VB.TextBox txtProvider 
      Height          =   300
      Left            =   960
      TabIndex        =   4
      Top             =   960
      Width           =   3495
   End
   Begin VB.CommandButton cmdFile 
      Caption         =   "…"
      Height          =   300
      Left            =   4440
      TabIndex        =   2
      Top             =   480
      Width           =   280
   End
   Begin VB.TextBox txtFile 
      Enabled         =   0   'False
      Height          =   300
      Left            =   960
      TabIndex        =   1
      Top             =   480
      Width           =   3495
   End
   Begin VB.ComboBox cboIOType 
      Height          =   300
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   960
      Width           =   3780
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfList 
      Height          =   2565
      Left            =   240
      TabIndex        =   8
      Top             =   1560
      Width           =   4935
      _cx             =   8705
      _cy             =   4524
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
      BackColorSel    =   4227072
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
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
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
      FormatString    =   $"frmPurchaseImportFile.frx":13916
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
      Editable        =   0
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
   Begin VSFlex8Ctl.VSFlexGrid vsfError 
      Height          =   765
      Left            =   240
      TabIndex        =   9
      Top             =   5040
      Width           =   4935
      _cx             =   8705
      _cy             =   1349
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
      ForeColorSel    =   16711680
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
      FormatString    =   $"frmPurchaseImportFile.frx":1398B
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
      Editable        =   0
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
   Begin XtremeCommandBars.ImageManager imgPicture 
      Left            =   840
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmPurchaseImportFile.frx":13A00
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin VB.Label lblProvider 
      AutoSize        =   -1  'True
      Caption         =   "供应商(&P)"
      Height          =   180
      Left            =   120
      TabIndex        =   3
      Top             =   1005
      Width           =   810
   End
   Begin VB.Label lblFile 
      AutoSize        =   -1  'True
      Caption         =   "文  件(&F)"
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   525
      Width           =   810
   End
End
Attribute VB_Name = "frmPurchaseImportFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MCONTOOLMODE As Integer = 100 '样本
Private Const MCONTOOLOUTPUT As Integer = 101   '导出文件
Private Const MCONTOOLCHECK As Integer = 102    '校验
Private Const MCONTOOLSAVE As Integer = 103 '保存
Private Const MCONTOOLEXIT As Integer = 104 '退出
Private Const MCONERROR As Integer = 105 '错误提示图标
Private Const MCONWARN   As Integer = 106   '警告图标
Private Const MCONYESCHECK As Integer = 107 '严格检查
Private Const mconNOCHECK As Integer = 108 '不严格检查
Private Const MCONTOOLCHECKCONDITION As Integer = 107 '检查条件

Private Mstr_Cols As String  'EXCEL样本列名
Private mblnResult As Boolean
Private mblnChange As Boolean
Private mlngModule As Long
Private mlngStockID As Long
Private mstrStock As String
Private mblnVirtualStock As Boolean '是否是虚拟库房，true-是 false-不是
Private mintUnit  As Integer                    '显示单位:0-散装单位,1-包装单位
Private mFMT As g_FmtString
'规则为导入方式/卫材编码|数量|成本价|成本金额|发票金额|数量*成本价=成本金额|发票金额=成本金额|表格成本价=HIS成本价|效期|灭菌日期|灭菌效期|生产日期|存储库房|虚拟库房|商品条码(0-不完全导入1-完全导入/0-提示1-禁止|....)
Private mbyt导入方式, mbyt卫材编码, mbyt数量, mbyt成本价, mbyt成本金额, mbyt发票金额, mbyt发票日期, mbytNumCost, mbytInvoiceCost, mbytExcelCost, mbyt效期, mbyt灭菌日期, mbyt灭菌效期, mbyt生产日期, mbyt存储库房, mbyt虚拟库房, mbyt商品条码 As Byte

'Private mobjXLS As Excel.Application
'Private mobjWB As Excel.Workbook
'Private mobjWS As Excel.Worksheet
Private mobjXLS As Object
Private mobjWB As Object
Private mobjWS As Object

Private Sub InitComandbar()
    '初始化工具栏
    Dim cbrControlMain As CommandBarControl
    Dim cbrToolBar As CommandBar
    Dim cbrToolPopup As CommandBarPopup
    
    'CommandBars
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    
    Me.cbsMain.VisualTheme = xtpThemeOffice2003

    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbsMain.EnableCustomization False
    Me.cbsMain.Icons = Me.imgPicture.Icons
    
    '工具栏定义
    Set cbrToolBar = Me.cbsMain.Add("工具栏", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched Or xtpFlagFloating Or xtpFlagAlignAny
        
    With cbrToolBar.Controls
'        Set cbrControlMain = .Add(xtpControlSplitButtonPopup, MCONTOOLMODE, "Excel样本")
'        cbrControlMain.CommandBar.Controls.Add xtpControlButton, MCONTOOLMODEEXCEL, "Excel样本"    '样本子菜单
'        cbrControlMain.CommandBar.Controls.Add xtpControlButton, MCONTOOLMODEXML, "XML样本"
'        cbrControlMain.Style = xtpButtonIconAndCaption  '同时显示图标和文字
        
'        Set cbrControlMain = .Add(xtpControlSplitButtonPopup, MCONTOOLOUTPUT, "导出Excel")
'        cbrControlMain.CommandBar.Controls.Add xtpControlButton, MCONTOOLOUTPUTEXCEL, "导出Excel"  '导出文件子菜单
'        cbrControlMain.CommandBar.Controls.Add xtpControlButton, MCONTOOLOUTPUTXML, "导出XML"
        Set cbrControlMain = .Add(xtpControlButton, MCONTOOLMODE, "Excel样本")
        cbrControlMain.Style = xtpButtonIconAndCaption  '同时显示图标和文字
        Set cbrControlMain = .Add(xtpControlButton, MCONTOOLOUTPUT, "导出Excel")
        cbrControlMain.Style = xtpButtonIconAndCaption  '同时显示图标和文字
        Set cbrControlMain = .Add(xtpControlButton, MCONTOOLCHECKCONDITION, "检查设置")
        cbrControlMain.Style = xtpButtonIconAndCaption  '同时显示图标和文字
        cbrControlMain.BeginGroup = True
        Set cbrControlMain = .Add(xtpControlButton, MCONTOOLCHECK, "校验")
        cbrControlMain.Style = xtpButtonIconAndCaption  '同时显示图标和文字
        cbrControlMain.BeginGroup = True
        Set cbrControlMain = .Add(xtpControlButton, MCONTOOLSAVE, "保存")
        cbrControlMain.Style = xtpButtonIconAndCaption  '同时显示图标和文字
        Set cbrControlMain = .Add(xtpControlButton, MCONTOOLEXIT, "退出")
        cbrControlMain.BeginGroup = True
        cbrControlMain.Style = xtpButtonIconAndCaption  '同时显示图标和文字
    End With
    cbsMain.Item(1).Delete
End Sub

Public Property Get Result() As Boolean
    Result = mblnResult
End Property

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.Id
        Case MCONTOOLMODE    '生成样本
            Call ProduceStyleBook
        Case MCONTOOLOUTPUT '导出文件
            Call OutPutFile
        Case MCONTOOLCHECK  '重新校验
            Call CheckData
        Case MCONTOOLSAVE   '保存
            Call SaveCard
        Case MCONTOOLCHECKCONDITION '条件设置
            frmPurchaseImportFileCondition.ShowMe Me, mlngModule
            If vsfList.Rows > 1 Then Call CheckData
        Case MCONTOOLEXIT    '退出
            Unload Me
    End Select
End Sub

Private Sub OutPutFile()
    '导出表格文件
    Dim strFileName As String
    Dim lngRow As Long
    Dim lngCol As Long
    Dim lng编码 As Long
    
    On Error GoTo ErrHandle
    Call InitExcel
    Set mobjWB = mobjXLS.Workbooks.Add
    Set mobjWS = mobjWB.ActiveSheet
    
    With vsfList
        If .Rows = 1 Then Exit Sub
        lng编码 = GetColumnPostation("卫材编码") + Asc("A") - 1 '走到这一步证明肯定有卫材编码列
        mobjWS.Range(Chr(lng编码) & "1:" & Chr(lng编码) & .Rows - 1).NumberFormatLocal = "@"
        For lngRow = 0 To .Rows - 1
            For lngCol = 1 To .Cols - 1
                If lngCol = 1 And lngRow <> 0 Then
                    mobjWS.cells(lngRow + 1, lngCol) = "'" & .TextMatrix(lngRow, lngCol)
                Else
                    mobjWS.cells(lngRow + 1, lngCol) = .TextMatrix(lngRow, lngCol)
                End If
            Next
        Next
    End With
    
    With dlgOpenFile
        .CancelError = True
        .FileName = ""
        .Filter = "*.xls|*.xls|*.xlsx|*.xlsx"
        .ShowSave
        strFileName = .FileName
        If Trim(strFileName) <> "" Then
            mobjWB.SaveAs strFileName
            mobjWB.Close
            Set mobjWS = Nothing
            Set mobjWB = Nothing
            mobjXLS.quit
            MsgBox "保存成功！", vbInformation, gstrSysName
        End If
    End With
    Exit Sub
    
ErrHandle:
    mobjWB.Close
    Set mobjWS = Nothing
    Set mobjWB = Nothing
    mobjXLS.quit
End Sub

Private Sub chkNoCheck_Click()
    If vsfList.Rows > 1 Then
        Call CheckData
    End If
End Sub

Private Sub chkYesCheck_Click()
    If vsfList.Rows > 1 Then
        Call CheckData
    End If
End Sub

Private Sub cmdFile_Click()
    On Error GoTo ErrHandle
    
    dlgOpenFile.FileName = ""
    dlgOpenFile.Filter = "*.xls|*.xls|*.xlsx|*.xlsx"
    dlgOpenFile.ShowOpen
    If dlgOpenFile.FileName <> "" Then
        txtFile.Text = dlgOpenFile.FileName
        If mlngModule = 1712 Then
            txtProvider.SetFocus
        ElseIf mlngModule = 1714 Then
            cboIOType.SetFocus
        End If
    Else
        GoTo ErrHandle
    End If
    If txtFile.Text <> "" Then
        Call ParseParameter
        DoEvents
        Call FS.ShowFlash("正在加载数据,请稍候 ...", Me)
        Me.MousePointer = vbHourglass
        
        ProCheck.Value = 0
        ProCheck.Visible = True
        Call InitExcel
        Call GetExcelData
        
        Me.MousePointer = vbDefault
        Call FS.StopFlash
        ProCheck.Visible = False
    End If
    Exit Sub
    
ErrHandle:
    Exit Sub
End Sub

Private Sub ParseParameter()
    '解析参数
    Dim i As Integer
    Dim arryPara As Variant
    Dim strPara As String
    
    If mlngModule = 1712 Then
        strPara = zlDatabase.GetPara("导入文件检查方式", glngSys, mlngModule, "0/0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0")
    Else
        strPara = zlDatabase.GetPara("导入文件检查方式", glngSys, mlngModule, "0/0|0|0|0|0|0|0|0|0|0|0|0|0|0|0|0")
    End If
    
    mbyt导入方式 = Mid(strPara, 1, 1)
    strPara = Mid(strPara, 3)
    arryPara = Split(strPara, "|")
    mbyt卫材编码 = arryPara(0)
    mbyt数量 = arryPara(1)
    mbyt成本价 = arryPara(2)
    mbyt成本金额 = arryPara(3)
    mbyt发票金额 = arryPara(4)
    mbyt发票日期 = arryPara(5)
    mbytNumCost = arryPara(6)
    mbytInvoiceCost = arryPara(7)
    mbytExcelCost = arryPara(8)
    mbyt效期 = arryPara(9)
    mbyt灭菌日期 = arryPara(10)
    mbyt灭菌效期 = arryPara(11)
    mbyt生产日期 = arryPara(12)
    mbyt存储库房 = arryPara(13)
    mbyt虚拟库房 = arryPara(14)
    mbyt商品条码 = arryPara(15)
End Sub

Private Sub ProduceStyleBook()
'生成导入外部文件的标准XLS文件样本
    Dim arrCols As Variant
    Dim i As Byte
    Dim blnFinished As Boolean
    Dim strFileName As String
    
    arrCols = Split(Mstr_Cols, ";")
    
    On Error GoTo ErrHandle
    Call InitExcel
    Set mobjWB = mobjXLS.Workbooks.Add
    Set mobjWS = mobjWB.ActiveSheet
    
    For i = LBound(arrCols) + 1 To UBound(arrCols)
        mobjWS.cells(1, i) = arrCols(i)
    Next
    
    With dlgOpenFile
        .FileName = ""
        .Filter = "Excel Files (*.xls)|*.xls"
        .ShowSave
        strFileName = .FileName
        If Trim(strFileName) <> "" Then
            mobjWB.SaveAs strFileName
            blnFinished = True
        Else
            strFileName = "False"
        End If
    End With
        
ErrHandle:
    mobjWB.Close
    Set mobjWS = Nothing
    Set mobjWB = Nothing
    If blnFinished Then
        MsgBox "标准文件样本已经生成！", vbInformation, gstrSysName
    ElseIf Trim(strFileName) <> "False" Then
        MsgBox "生成标准文件样本失败！", vbInformation, gstrSysName
    End If
End Sub

Private Sub cmdProvider_Click()
    If Select供应商(Me, txtProvider, "") Then
        OS.PressKey vbKeyTab
    Else
        txtProvider.SetFocus
    End If
End Sub

Private Function CheckQualifications(ByVal intType As Integer, ByVal strInput As String) As Boolean
    '校验卫材，生产商，供应商信息和资质效期
    'intType：0－卫材；1－生产商；2－供应商
    'strInput：字符串时为名称；数字时为ID
    Dim rsTmp As ADODB.Recordset
    Dim strMsgInfo As String
    Dim strMsgDate As String
    Dim dateCurrent As Date
    Dim strMsg As String
    
    Dim intCheckType As Integer
    Dim arrColumn
    Dim strCheck As String
    Dim strCheck_卫材 As String
    Dim strCheck_生产商 As String
    Dim strCheck_供应商 As String
    Dim n As Integer
    Dim strTmp As String
    
    On Error GoTo ErrHandle
    If strInput = "" Then
        CheckQualifications = True
        Exit Function
    End If
        
    '资质校验项目和方式的保存格式：校验方式|类别1,项目1,是否校验;类别1,项目2,是否校验;类别2,项目1,是否校验;类别2,项目2....
    strCheck = zlDatabase.GetPara("资质校验", glngSys, mlngModule, "")
    
    '保存的参数格式不正确时退出
    If InStr(1, strCheck, "|") = 0 Then
        CheckQualifications = True
        Exit Function
    End If

    '取校验方式：0-不检查；1－提醒；2－禁止
    intCheckType = Val(Mid(strCheck, 1, InStr(1, strCheck, "|") - 1))
    
    '不检查时退出
    If intCheckType = 0 Then
        CheckQualifications = True
        Exit Function
    End If

    '取校验内容：
    strCheck = Mid(strCheck, InStr(1, strCheck, "|") + 1)

    If strCheck = "" Then
        CheckQualifications = True
        Exit Function
    End If

    '分别取卫材，生产商，供应商需要校验的内容
    strCheck = strCheck & ";"
    arrColumn = Split(strCheck, ";")
    For n = 0 To UBound(arrColumn)
        If arrColumn(n) <> "" Then
            If Split(arrColumn(n), ",")(0) = "卫材" And Split(arrColumn(n), ",")(2) = 1 Then
                strCheck_卫材 = IIf(strCheck_卫材 = "", "", strCheck_卫材 & ";") & Split(arrColumn(n), ",")(1)
            End If

            If Split(arrColumn(n), ",")(0) = "卫材生产商" And Split(arrColumn(n), ",")(2) = 1 Then
                strCheck_生产商 = IIf(strCheck_生产商 = "", "", strCheck_生产商 & ";") & Split(arrColumn(n), ",")(1)
            End If

            If Split(arrColumn(n), ",")(0) = "卫材供应商" And Split(arrColumn(n), ",")(2) = 1 Then
                strCheck_供应商 = IIf(strCheck_供应商 = "", "", strCheck_供应商 & ";") & Split(arrColumn(n), ",")(1)
            End If
        End If
    Next
    
    '无校验内容时退出
    If (intType = 0 And strCheck_卫材 = "") Or (intType = 1 And strCheck_生产商 = "") Or (intType = 2 And strCheck_供应商 = "") Then
        CheckQualifications = True
        Exit Function
    End If
    
    dateCurrent = CDate(Format(sys.Currentdate, "yyyy-mm-dd"))
    
    '卫材
    If intType = 0 Then
        gstrSQL = "Select ('[' || B.编码 || ']' || B.名称) AS 卫材信息, A.许可证号, A.许可证有效期 " & _
            " From 收费项目目录 B,材料特性 A " & _
            " Where B.ID = A.材料ID And A.材料ID = [1] "
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "校验卫材资质", Val(strInput))
        
        If Not rsTmp.EOF Then
            If zlStr.Nvl(rsTmp!许可证号) = "" And InStr(strCheck_卫材, "许可证号") > 0 Then
                strTmp = rsTmp!卫材信息 & "：" & "无许可证号"
            End If
            
            If zlStr.Nvl(rsTmp!许可证有效期) <> "" Then
                If DateDiff("d", rsTmp!许可证有效期, dateCurrent) > 0 And InStr(strCheck_卫材, "许可证有效期") > 0 Then
                    strTmp = IIf(strTmp = "", rsTmp!卫材信息 & "：", strTmp & ",") & "许可证已过期"
                End If
            End If
        End If
    End If
    
    '生产商
    If intType = 1 Then
        gstrSQL = "Select ('[' || A.编码 || ']' || A.名称) AS 生产商, A.生产企业许可证, A.生产企业许可证效期,a.经营许可证, a.经营许可证效期, a.企业法人执照, a.企业法人执照效期 " & _
                        " From 材料生产商 A " & _
                        " Where A.名称 = [1] "
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "校验卫材资质", strInput)
        
        If Not rsTmp.EOF Then
            If zlStr.Nvl(rsTmp!生产企业许可证) = "" And InStr(strCheck_生产商 & ";", "生产企业许可证" & ";") > 0 Then
                strTmp = rsTmp!生产商 & "：" & "无生产企业许可证"
            End If
            
            If zlStr.Nvl(rsTmp!生产企业许可证效期) <> "" Then
                If DateDiff("d", rsTmp!生产企业许可证效期, dateCurrent) > 0 And InStr(strCheck_生产商, "生产企业许可证效期") > 0 Then
                    strTmp = IIf(strTmp = "", rsTmp!生产商 & "：", strTmp & ",") & "生产企业许可证已过期"
                End If
            End If
        End If
        
        If Not rsTmp.EOF Then
            If zlStr.Nvl(rsTmp!经营许可证) = "" And InStr(strCheck_生产商 & ";", "经营许可证" & ";") > 0 Then
                strTmp = rsTmp!生产商 & "：" & "无经营许可证"
            End If
            
            If zlStr.Nvl(rsTmp!经营许可证效期) <> "" Then
                If DateDiff("d", rsTmp!生产企业许可证效期, dateCurrent) > 0 And InStr(strCheck_生产商 & ";", "经营许可证效期" & ";") > 0 Then
                    strTmp = IIf(strTmp = "", rsTmp!生产商 & "：", strTmp & ",") & "经营许可证已过期"
                End If
            End If
        End If
        If Not rsTmp.EOF Then
            If zlStr.Nvl(rsTmp!企业法人执照) = "" And InStr(strCheck_生产商 & ";", "企业法人执照" & ";") > 0 Then
                strTmp = rsTmp!生产商 & "：" & "无企业法人执照"
            End If
            
            If zlStr.Nvl(rsTmp!企业法人执照效期) <> "" Then
                If DateDiff("d", rsTmp!生产企业许可证效期, dateCurrent) > 0 And InStr(strCheck_生产商 & ";", "企业法人执照效期" & ";") > 0 Then
                    strTmp = IIf(strTmp = "", rsTmp!生产商 & "：", strTmp & ",") & "企业法人执照已过期"
                End If
            End If
        End If
    End If
    
    '供应商
    If intType = 2 Then
        gstrSQL = "Select ('[' || 编码 || ']' || 名称) AS 供应商, 税务登记号, 许可证号, 执照号, 授权号, 质量认证号, 质量认证日期, 药监局备案号, 药监局备案日期, 许可证效期, 执照效期, 授权期 " & _
            " From 供应商 " & _
            " Where (撤档时间 Is Null Or 撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) And ID = [1] "
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "供应商信息", Val(strInput))
        
        strTmp = ""
        
        If Not rsTmp.EOF Then
            If zlStr.Nvl(rsTmp!税务登记号) = "" And InStr(strCheck_供应商, "税务登记号") > 0 Then
                strTmp = rsTmp!供应商 & "：" & "无税务登记号"
            End If
            
            If zlStr.Nvl(rsTmp!许可证号) = "" And InStr(strCheck_供应商, "许可证号") > 0 Then
                strTmp = IIf(strTmp = "", rsTmp!供应商 & "：", strTmp & ",") & "无许可证号"
            End If
            
            If zlStr.Nvl(rsTmp!执照号) = "" And InStr(strCheck_供应商, "执照号") > 0 Then
                strTmp = IIf(strTmp = "", rsTmp!供应商 & "：", strTmp & ",") & "无执照号"
            End If
            
            If zlStr.Nvl(rsTmp!授权号) = "" And InStr(strCheck_供应商, "授权号") > 0 Then
                strTmp = IIf(strTmp = "", rsTmp!供应商 & "：", strTmp & ",") & "无授权号"
            End If
            
            If zlStr.Nvl(rsTmp!质量认证号) = "" And InStr(strCheck_供应商, "质量认证号") > 0 Then
                strTmp = IIf(strTmp = "", rsTmp!供应商 & "：", strTmp & ",") & "无质量认证号"
            End If
            
            If zlStr.Nvl(rsTmp!质量认证日期) <> "" Then
                If DateDiff("d", rsTmp!质量认证日期, dateCurrent) > 0 And InStr(strCheck_供应商, "质量认证日期") > 0 Then
                    strTmp = IIf(strTmp = "", rsTmp!供应商 & "：", strTmp & ",") & "质量认证号已过期"
                End If
            End If
            
            If zlStr.Nvl(rsTmp!药监局备案号) = "" And InStr(strCheck_供应商, "药监局备案号") > 0 Then
                strTmp = IIf(strTmp = "", rsTmp!供应商 & "：", strTmp & ",") & "无药监局备案号"
            End If
            
            If zlStr.Nvl(rsTmp!药监局备案日期) <> "" Then
                If DateDiff("d", rsTmp!药监局备案日期, dateCurrent) > 0 And InStr(strCheck_供应商, "药监局备案日期") > 0 Then
                    strTmp = IIf(strTmp = "", rsTmp!供应商 & "：", strTmp & ",") & "药监局备案号已过期"
                End If
            End If
            
            If zlStr.Nvl(rsTmp!许可证效期) <> "" Then
                If DateDiff("d", rsTmp!许可证效期, dateCurrent) > 0 And InStr(strCheck_供应商, "许可证效期") > 0 Then
                    strTmp = IIf(strTmp = "", rsTmp!供应商 & "：", strTmp & ",") & "许可证已过期"
                End If
            End If
            
            If zlStr.Nvl(rsTmp!执照效期) <> "" Then
                If DateDiff("d", rsTmp!执照效期, dateCurrent) > 0 And InStr(strCheck_供应商, "执照效期") > 0 Then
                    strTmp = IIf(strTmp = "", rsTmp!供应商 & "：", strTmp & ",") & "执照已过期"
                End If
            End If
            
            If zlStr.Nvl(rsTmp!授权期) <> "" Then
                If DateDiff("d", rsTmp!执照效期, dateCurrent) > 0 And InStr(strCheck_供应商, "授权期") > 0 Then
                    strTmp = IIf(strTmp = "", rsTmp!供应商 & "：", strTmp & ",") & "授权已过期"
                End If
            End If
        End If
    End If
    
    '提示或禁止
    If strTmp <> "" Then
        If intCheckType = 1 Then
            If MsgBox("未通过资质校验，是否继续？" & vbCrLf & strTmp, vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
                CheckQualifications = True
                Exit Function
            Else
                Exit Function
            End If
        ElseIf intCheckType = 2 Then
            MsgBox "未通过资质校验，不能入库！" & vbCrLf & strTmp, vbOKOnly, gstrSysName
            Exit Function
        End If
    End If
    
    CheckQualifications = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function EntryPort(ByVal lngModule As Long, ByVal strStockInfo As String)
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    mlngModule = lngModule
    mlngStockID = Val(Split(strStockInfo, ";")(0))
    mstrStock = Split(strStockInfo, ";")(1)
    Caption = "导入外部文件(" & mstrStock & ")"
    
    Select Case mlngModule
    Case 1712
        Mstr_Cols = ";卫材编码;卫材名称;规格;产地;批号;生产日期;效期;灭菌日期;灭菌效期;数量;单位;成本价;成本金额;发票号;发票日期;发票金额;商品码;"
        lblProvider.Caption = "供应商(&P)"
        cboIOType.Visible = False
    Case 1714
        Mstr_Cols = ";卫材编码;卫材名称;规格;产地;批号;生产日期;效期;灭菌日期;灭菌效期;数量;单位;成本价;成本金额;商品码;"
        lblProvider.Caption = "入出类(&I)"
        txtProvider.Visible = False
        cmdProvider.Visible = False
        cboIOType.Top = txtProvider.Top
        gstrSQL = "Select b.Id, b.名称 From 药品单据性质 A, 药品入出类别 B Where a.类别id = b.Id And a.单据 = 32"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        cboIOType.Clear
        Do While Not rsTmp.EOF
            cboIOType.AddItem rsTmp!名称
            cboIOType.ItemData(cboIOType.NewIndex) = rsTmp!Id
            rsTmp.MoveNext
        Loop
        If rsTmp.RecordCount > 0 Then cboIOType.ListIndex = 0
    End Select
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SaveCard()
    Dim lngRow As Long
    Dim lngCol As Long
    
    On Error GoTo ErrHandle
    
    If mlngModule = 1712 Then
        If Val(txtProvider.Tag) = 0 Then
            MsgBox "未选择供应商！", vbInformation, gstrSysName
            txtProvider.SetFocus
            Exit Sub
        End If
    ElseIf mlngModule = 1714 Then
        If cboIOType.ListIndex < 0 Then
            MsgBox "未选择入出类！", vbInformation, gstrSysName
            cboIOType.SetFocus
            Exit Sub
        End If
    End If
    If vsfList.Rows = 1 Then Exit Sub
    vsfError.Rows = 1
    Call CheckData '保存时检查数据
    With vsfError
        If .Rows > 1 Then
            For lngRow = 1 To .Rows - 1
                If mbyt导入方式 = 1 Then
                    If .Cell(flexcpPicture, lngRow, 0, lngRow, 0) = imgError.ListImages(2).Picture Or .Cell(flexcpPicture, lngRow, 0, lngRow, 0) = imgError.ListImages(1).Picture Then
                        MsgBox "完全导入方式下，不能存在任何不合格的数据，请修正！", vbInformation, gstrSysName
                        Exit Sub
                    End If
                Else
                    If .Cell(flexcpPicture, lngRow, 0, lngRow, 0) = imgError.ListImages(2).Picture Or .Cell(flexcpPicture, lngRow, 0, lngRow, 0) = imgError.ListImages(1).Picture Then
                        If MsgBox("还存在不合格数据，是否继续？", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                            Exit Sub
                        Else
                            Exit For
                        End If
                    End If
                End If
            Next
        End If
    End With
        
    '导入
    Call ImportData
    Exit Sub
    
ErrHandle:
    If Not mobjWB Is Nothing Then
        mobjWB.Close
    End If
    Set mobjWB = Nothing
    Call ErrCenter
    Call SaveErrLog
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub Form_Load()
    Dim rsTemp As ADODB.Recordset
    Dim cbrControl As CommandBarControl
    
    On Error GoTo ErrHandle
    
    Me.Height = 600 * 15
    Me.Width = 800 * 15
    
    Call InitComandbar
    Call InitControlPosition
    Call InitVSF
    ProCheck.Value = 100
    mintUnit = 1
    With mFMT
        .FM_成本价 = GetFmtString(mintUnit, g_成本价)
        .FM_金额 = GetFmtString(mintUnit, g_金额)
        .FM_零售价 = GetFmtString(mintUnit, g_售价)
        .FM_数量 = GetFmtString(mintUnit, g_数量)
        .FM_散装零售价 = GetFmtString(0, g_售价)
    End With
    gstrSQL = "Select 1 From 部门性质说明 Where 部门id = [1] And 工作性质 = '虚拟库房' And Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "库房性质查询", mlngStockID)
    If rsTemp.RecordCount > 0 Then mblnVirtualStock = True
        
    If vsfList.Rows = 1 Then
        Set cbrControl = Me.cbsMain(1).Controls.Find(xtpControlButton, MCONTOOLOUTPUT, , True)
        cbrControl.Enabled = False
        Set cbrControl = Me.cbsMain(1).Controls.Find(xtpControlButton, MCONTOOLCHECK, , True)
        cbrControl.Enabled = False
        Set cbrControl = Me.cbsMain(1).Controls.Find(xtpControlButton, MCONTOOLSAVE, , True)
        cbrControl.Enabled = False
    End If
    
    Exit Sub

ErrHandle:
    Call ErrCenter
    Call SaveErrLog
End Sub

Private Sub InitVSF()
    '初始化表格控件
    With vsfList
        .Rows = 1
        .Cols = 17
        .Editable = flexEDKbdMouse
        .ExplorerBar = flexExNone   '列不支持排序和拖动
    End With
    
    With vsfError
        .Rows = 1
        .Cols = 4
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "错误位置"
        .ColWidth(1) = 2000
        .TextMatrix(0, 2) = "错误类型"
        .ColWidth(2) = 1500
        .TextMatrix(0, 3) = "错误原因"
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
'        .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = True
        .ColWidth(0) = 300
        .ExtendLastCol = True '最后一列填充满
        .ExplorerBar = flexExNone   '列不支持排序和拖动
    End With
End Sub

Private Sub InitExcel()
    '初始化Excel表格
    Set mobjXLS = CreateObject("Excel.Application")
    mobjXLS.DisplayAlerts = False
End Sub

Private Function GetExcelData()
    '获取excel表格数据，并将其显示到界面表格中
    '返回true-有错误 返回false-没有错误
    Dim strFileColumn As String '文件中列名称
    Dim lngRow As Long
    Dim lngCol As Long
    Dim intNum As Integer   '数量列
    Dim intCost As Integer  '成本价列
    Dim dbl成本金额 As Double
    Dim bln成本金额 As Boolean
    Dim dbl发票金额 As Double
    Dim bln发票金额 As Double
    Dim blnNotNullRow As Boolean   '检查改行不是空行
    Dim bln生产日期 As Boolean
    Dim bln效期 As Boolean
    Dim bln灭菌日期 As Boolean
    Dim bln灭菌效期 As Boolean
    Dim bln发票日期 As Boolean
    
    Dim str生产日期 As String
    Dim str效期 As String
    Dim str灭菌日期 As String
    Dim str灭菌效期 As String
    Dim str发票日期 As String
    
    On Error GoTo ErrHandle
    If txtFile.Text = "" Then Exit Function
    
    Set mobjWB = mobjXLS.Workbooks.Open(txtFile.Text)
    Set mobjWS = mobjWB.Sheets(1)
    If mobjWS Is Nothing Then Exit Function
    
    With mobjWS.UsedRange
        '列名和列顺序检查
        For lngCol = 1 To .Columns.count
            strFileColumn = strFileColumn & ";" & .cells(1, lngCol)
        Next
        For lngCol = 1 To UBound(Split(strFileColumn, ";"))
            If InStr(1, Mstr_Cols, ";" & Split(strFileColumn, ";")(lngCol) & ";") = 0 Then
                vsfError.Rows = vsfError.Rows + 1
                vsfError.TextMatrix(vsfError.Rows - 1, 1) = "1行"
                vsfError.TextMatrix(vsfError.Rows - 1, 2) = "列头错误"
                vsfError.TextMatrix(vsfError.Rows - 1, 3) = "导入文件列头【" & Split(strFileColumn, ";")(lngCol) & "】不存在，正确列头应该是【" & Mstr_Cols & "】请修正要导入的Excel文件！"
                GetExcelData = True
            End If
        Next
        If mbyt导入方式 = 1 Then
            If mbyt卫材编码 = 1 Then
                If InStr(1, strFileColumn, ";卫材编码;") = 0 Then
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = "1行"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "列头错误"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【卫材编码】列不存在，请修正要导入的Excel文件！"
                    GetExcelData = True
                End If
            End If
            If mbyt数量 = 1 Then
                If InStr(1, strFileColumn, ";数量;") = 0 Then
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = "1行"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "列头错误"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【数量】列不存在，请修正要导入的Excel文件！"
                    GetExcelData = True
                End If
            End If
            If mbyt成本价 = 1 Then
                If InStr(1, strFileColumn, ";成本价;") = 0 Then
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = "1行"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "列头错误"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【成本价】列不存在，请修正要导入的Excel文件！"
                    GetExcelData = True
                End If
            End If
            If mbyt成本金额 = 1 Then
                If InStr(1, strFileColumn, ";成本金额;") = 0 Then
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = "1行"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "列头错误"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【成本金额】列不存在，请修正要导入的Excel文件！"
                    GetExcelData = True
                End If
            End If
        End If
        '上述检查不能通过则需要修改导入文件，因此退出检查
        If GetExcelData = True Then
            Exit Function
        End If
        '往表格中填充数据
        vsfList.Redraw = flexRDNone
        vsfList.Cols = .Columns.count + 1
        vsfList.Rows = 1
        For lngRow = 1 To .Rows.count
            vsfList.Rows = vsfList.Rows + 1
            For lngCol = 1 To .Columns.count
                vsfList.TextMatrix(lngRow - 1, lngCol) = .cells(lngRow, lngCol)
                If lngRow = 1 Then
                    vsfList.ColKey(lngCol) = .cells(lngRow, lngCol)
                End If
            Next
        Next
        '将空行删除
        For lngRow = vsfList.Rows - 1 To 1 Step -1
            blnNotNullRow = True
            For lngCol = 1 To vsfList.Cols - 1
                If vsfList.TextMatrix(lngRow, lngCol) <> "" Then
                    blnNotNullRow = False
                End If
            Next
            '如果是空行将其删除
            If blnNotNullRow = True Then vsfList.RemoveItem lngRow
        Next
        Set mobjWS = Nothing
        Set mobjWB = Nothing
        mobjXLS.quit
        Set mobjXLS = Nothing
        With vsfList
            bln成本金额 = IIf(GetColumnPostation("成本金额") > 0, True, False)
            bln发票金额 = IIf(GetColumnPostation("发票金额") > 0, True, False)
            bln生产日期 = IIf(GetColumnPostation("生产日期") > 0, True, False)
            bln效期 = IIf(GetColumnPostation("效期") > 0, True, False)
            bln灭菌日期 = IIf(GetColumnPostation("灭菌日期") > 0, True, False)
            bln灭菌效期 = IIf(GetColumnPostation("灭菌效期") > 0, True, False)
            bln发票日期 = IIf(GetColumnPostation("发票日期") > 0, True, False)
            For lngRow = 1 To .Rows - 1
                If bln生产日期 = True Then
                    str生产日期 = FormatDate(.TextMatrix(lngRow, .ColIndex("生产日期")))
                    .TextMatrix(lngRow, .ColIndex("生产日期")) = IIf(str生产日期 = "", .TextMatrix(lngRow, .ColIndex("生产日期")), str生产日期)
                End If
                If bln效期 = True Then
                    str效期 = FormatDate(.TextMatrix(lngRow, .ColIndex("效期")))
                    .TextMatrix(lngRow, .ColIndex("效期")) = IIf(str效期 = "", .TextMatrix(lngRow, .ColIndex("效期")), str效期)
                End If
                If bln灭菌日期 = True Then
                    str灭菌日期 = FormatDate(.TextMatrix(lngRow, .ColIndex("灭菌日期")))
                    .TextMatrix(lngRow, .ColIndex("灭菌日期")) = IIf(str灭菌日期 = "", .TextMatrix(lngRow, .ColIndex("灭菌日期")), str灭菌日期)
                End If
                If bln灭菌效期 = True Then
                    str灭菌效期 = FormatDate(.TextMatrix(lngRow, .ColIndex("灭菌效期")))
                    .TextMatrix(lngRow, .ColIndex("灭菌效期")) = IIf(str灭菌效期 = "", .TextMatrix(lngRow, .ColIndex("灭菌效期")), str灭菌效期)
                End If
                If bln发票日期 = True Then
                    str发票日期 = FormatDate(.TextMatrix(lngRow, .ColIndex("发票日期")))
                    .TextMatrix(lngRow, .ColIndex("发票日期")) = IIf(str发票日期 = "", .TextMatrix(lngRow, .ColIndex("发票日期")), FormatDate(str发票日期))
                End If
                If bln成本金额 = True Then
                    dbl成本金额 = dbl成本金额 + Val(.TextMatrix(lngRow, .ColIndex("成本金额")))
                End If
                If bln发票金额 = True Then
                    dbl发票金额 = dbl发票金额 + Val(.TextMatrix(lngRow, .ColIndex("发票金额")))
                End If
            Next
            lblCollect.Caption = "成本金额：" & Format(dbl成本金额, mFMT.FM_金额) & "元          发票金额：" & Format(dbl发票金额, mFMT.FM_金额) & "元"
        End With
        Call SetColumn
        Call CheckData
        
        vsfList.Redraw = flexRDDirect
    End With
    Exit Function
    
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function FormatDate(ByVal strDate As String) As String
    '功能：格式化日期，返回用分号(-)分隔的日期格式
    '返回格式化后的值，如果为空则说明不是日期格式，不为空说明是日期格式
    Dim strYear, strMonth, strDay As String
    
    If LenB(StrConv(strDate, vbFromUnicode)) >= 8 Then
        If InStr(1, strDate, ".") > 0 Or InStr(1, strDate, "/") > 0 Or InStr(1, strDate, "-") > 0 Then
            strDate = Replace(strDate, ".", "")
            strDate = Replace(strDate, "/", "")
            strDate = Replace(strDate, "-", "")
        End If
        strYear = Mid(strDate, 1, 4)
        If LenB(StrConv(strDate, vbFromUnicode)) < 8 Then
            strMonth = Mid(strDate, 5, 1)
        Else
            strMonth = Mid(strDate, 5, 2)
        End If
        If LenB(StrConv(strDate, vbFromUnicode)) < 8 Then
            strDay = Mid(strDate, 6, 1)
        Else
            strDay = Mid(strDate, 7, 2)
        End If
        If IsNumeric(strYear) = True And IsNumeric(strMonth) = True And IsNumeric(strDay) = True Then
            FormatDate = strYear & "-" & strMonth & "-" & strDay
        End If
    Else
        FormatDate = ""
    End If
End Function

Private Function GetColumnPostation(ByVal strColumn As String) As Integer
    '获取列位置和判断是否存在
    '参数 strcolumn-传入的列名
    '返回值 :返回传入列位置 0-没有找到 >0找到了
    Dim lngRow As Long
    Dim lngCol As Long
    
    With vsfList
        For lngCol = 1 To .Cols - 1
            If .TextMatrix(0, lngCol) = strColumn Then
                GetColumnPostation = lngCol
                Exit Function
            End If
        Next
        GetColumnPostation = 0
    End With
End Function

Private Sub CheckData()
    '检查数据合法性
    Dim lngRow As Long
    Dim lngCol As Long
    Dim rsTemp As ADODB.Recordset
    Dim dbl成本价 As Double
    Dim dbl数量 As Double
    Dim dbl成本金额 As Double
    Dim dbl发票金额 As Double
    Dim cbrControl As CommandBarControl
    Dim strTemp As String
    
    On Error GoTo ErrHandle
    
    Call ParseParameter
    vsfError.Rows = 1
    With vsfList
        .Cell(flexcpForeColor, 0, 0, .Rows - 1, .Cols - 1) = vbBlack '先设置成黑色
        ProCheck.Value = 0
        For lngRow = 1 To .Rows - 1
            .TextMatrix(lngRow, 0) = lngRow '添加行标
            '卫材编码
            If GetColumnPostation("卫材编码") > 0 Then
                If .TextMatrix(lngRow, .ColIndex("卫材编码")) = "" Then
                    .Cell(flexcpForeColor, lngRow, .ColIndex("卫材编码"), lngRow, .ColIndex("卫材编码")) = vbRed
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt卫材编码 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【卫材编码】列"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "空值错误"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【卫材编码】列为空，请修正！"
                Else
                    gstrSQL = "Select 1 From 收费项目目录 Where 类别 = '4' And 编码 =[1] And Rownum < 2"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "编码验证", .TextMatrix(lngRow, .ColIndex("卫材编码")))
                    If rsTemp.RecordCount = 0 Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("卫材编码"), lngRow, .ColIndex("卫材编码")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt卫材编码 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【卫材编码】列"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "数值错误"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【卫材编码】列编码不存在，请修正！"
                    End If
                End If
            End If
            '卫材名称
            If GetColumnPostation("卫材名称") > 0 Then
                .TextMatrix(lngRow, .ColIndex("卫材名称")) = Trim(.TextMatrix(lngRow, .ColIndex("卫材名称")))
            End If
            '规格
            If GetColumnPostation("规格") > 0 Then
                .TextMatrix(lngRow, .ColIndex("规格")) = Trim(.TextMatrix(lngRow, .ColIndex("规格")))
            End If
            '数量
            If .TextMatrix(lngRow, .ColIndex("数量")) = "" Then
                .Cell(flexcpForeColor, lngRow, .ColIndex("数量"), lngRow, .ColIndex("数量")) = vbRed
                vsfError.Rows = vsfError.Rows + 1
                vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt数量 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【数量】列"
                vsfError.TextMatrix(vsfError.Rows - 1, 2) = "空值错误"
                vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【数量】列是必填项且大于零，请修正！"
            Else
                If Not IsNumeric(.TextMatrix(lngRow, .ColIndex("数量"))) Then
                    .Cell(flexcpForeColor, lngRow, .ColIndex("数量"), lngRow, .ColIndex("数量")) = vbRed
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt数量 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【数量】列"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "格式错误"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【数量】列应为数字类型，请修正！"
                Else
                    .TextMatrix(lngRow, .ColIndex("数量")) = Format(.TextMatrix(lngRow, .ColIndex("数量")), mFMT.FM_数量)
                    If Val(.TextMatrix(lngRow, .ColIndex("数量"))) > 9999999999# Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("数量"), lngRow, .ColIndex("数量")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt数量 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【数量】列"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "数据错误"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【数量】列大于9999999999#，请修正！"
                    End If
                End If
            End If
            '成本价
            If .TextMatrix(lngRow, .ColIndex("成本价")) = "" Then
                .Cell(flexcpForeColor, lngRow, .ColIndex("成本价"), lngRow, .ColIndex("成本价")) = vbRed
                vsfError.Rows = vsfError.Rows + 1
                vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt成本价 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【成本价】列"
                vsfError.TextMatrix(vsfError.Rows - 1, 2) = "空值提示"
                vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【成本价】列为零了！"
            Else
                If Not IsNumeric(.TextMatrix(lngRow, .ColIndex("成本价"))) Then
                    .Cell(flexcpForeColor, lngRow, .ColIndex("成本价"), lngRow, .ColIndex("成本价")) = vbRed
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt成本价 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【成本价】列"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "格式错误"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【成本价】列应为数字类型，请修正！"
                Else
                    .TextMatrix(lngRow, .ColIndex("成本价")) = Format(.TextMatrix(lngRow, .ColIndex("成本价")), mFMT.FM_成本价)
                    If Val(.TextMatrix(lngRow, .ColIndex("成本价"))) > 999999999 Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("成本价"), lngRow, .ColIndex("成本价")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt成本价 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【成本价】列"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "数据错误"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【成本价】列大于999999999，请修正！"
                    End If
                End If
            End If
            '成本金额
            If .TextMatrix(lngRow, .ColIndex("成本金额")) = "" Then
                .Cell(flexcpForeColor, lngRow, .ColIndex("成本金额"), lngRow, .ColIndex("成本金额")) = vbRed
                vsfError.Rows = vsfError.Rows + 1
                vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt成本金额 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【成本金额】列"
                vsfError.TextMatrix(vsfError.Rows - 1, 2) = "空值提示"
                vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【成本金额】为零了！"
            Else
                If Not IsNumeric(.TextMatrix(lngRow, .ColIndex("成本金额"))) Then
                    .Cell(flexcpForeColor, lngRow, .ColIndex("成本金额"), lngRow, .ColIndex("成本金额")) = vbRed
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt成本金额 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【成本金额】列"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "格式错误"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【成本金额】列应为数字类型，请修正！"
                Else
                    .TextMatrix(lngRow, .ColIndex("成本金额")) = Format(.TextMatrix(lngRow, .ColIndex("成本金额")), mFMT.FM_金额)
                    If Val(.TextMatrix(lngRow, .ColIndex("成本金额"))) > 9999999999# Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("成本金额"), lngRow, .ColIndex("成本金额")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt成本金额 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【成本金额】列"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "格式错误"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【成本金额】列大于9999999999，请修正！"
                    End If
                End If
            End If
            '发票金额
            If GetColumnPostation("发票金额") > 0 Then
                If .TextMatrix(lngRow, .ColIndex("发票金额")) = "" Then
                    .Cell(flexcpForeColor, lngRow, .ColIndex("发票金额"), lngRow, .ColIndex("发票金额")) = vbRed
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt发票金额 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【发票金额】列"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "空值提示"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【发票金额】为零了！"
                Else
                    If Not IsNumeric(.TextMatrix(lngRow, .ColIndex("发票金额"))) Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("发票金额"), lngRow, .ColIndex("发票金额")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt发票金额 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【发票金额】列"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "格式错误"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【发票金额】列应为数字类型，请修正！"
                    Else
                        .TextMatrix(lngRow, .ColIndex("发票金额")) = Format(.TextMatrix(lngRow, .ColIndex("发票金额")), mFMT.FM_金额)
                        If Val(.TextMatrix(lngRow, .ColIndex("发票金额"))) > 999999999999# Then
                            .Cell(flexcpForeColor, lngRow, .ColIndex("发票金额"), lngRow, .ColIndex("发票金额")) = vbRed
                            vsfError.Rows = vsfError.Rows + 1
                            vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt发票金额 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                            vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【发票金额】列"
                            vsfError.TextMatrix(vsfError.Rows - 1, 2) = "数据错误"
                            vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【发票金额】列大于999999999999，请修正！"
                        End If
                    End If
                End If
            End If
            '数量*成本价=成本金额
            dbl成本价 = Format(Val(.TextMatrix(lngRow, .ColIndex("成本价"))), mFMT.FM_成本价)
            dbl数量 = Format(Val(.TextMatrix(lngRow, .ColIndex("数量"))), mFMT.FM_数量)
            dbl成本金额 = Format(Val(.TextMatrix(lngRow, .ColIndex("成本金额"))), mFMT.FM_金额)
            If dbl成本价 * dbl数量 <> dbl成本金额 Then
                .Cell(flexcpForeColor, lngRow, .ColIndex("成本金额"), lngRow, .ColIndex("成本金额")) = vbRed
                vsfError.Rows = vsfError.Rows + 1
                vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbytNumCost = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【成本金额】列"
                vsfError.TextMatrix(vsfError.Rows - 1, 2) = "数据错误"
                vsfError.TextMatrix(vsfError.Rows - 1, 3) = "成本价*数量<>成本金额，请修正！"
            End If
            '发票金额=成本金额
            If GetColumnPostation("发票金额") > 0 Then
                dbl发票金额 = Format(Val(.TextMatrix(lngRow, .ColIndex("发票金额"))), mFMT.FM_金额)
                If dbl发票金额 <> dbl成本金额 Then
                    .Cell(flexcpForeColor, lngRow, .ColIndex("发票金额"), lngRow, .ColIndex("发票金额")) = vbRed
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbytInvoiceCost = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【发票金额】列"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "数据提示"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "发票金额<>成本金额，请检查！"
                End If
            End If
            '文件成本价=HIS成本价
            gstrSQL = "Select Nvl(c.成本价, 0) 成本价, Nvl(c.换算系数, 1) As 换算系数" & vbNewLine & _
                    "From 收费项目目录 B, 材料特性 C" & vbNewLine & _
                    "Where b.Id = c.材料id And b.编码 = [1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "规格表成本价", .TextMatrix(lngRow, .ColIndex("卫材编码")))
            If rsTemp.RecordCount <> 0 Then
                If Format(Val(.TextMatrix(lngRow, .ColIndex("成本价"))), mFMT.FM_成本价) <> Format(rsTemp!成本价 * rsTemp!换算系数, mFMT.FM_成本价) Then
                    .Cell(flexcpForeColor, lngRow, .ColIndex("成本价"), lngRow, .ColIndex("成本价")) = vbRed
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbytExcelCost = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【成本价】列"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "数据提示"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "文档【成本价】" & Format(Val(.TextMatrix(lngRow, .ColIndex("成本价"))), mFMT.FM_成本价) & "与HIS系统规格中【成本价】" & Format(rsTemp!成本价 * rsTemp!换算系数, mFMT.FM_成本价) & "不等！"
                End If
            End If
            '效期
            If GetColumnPostation("效期") > 0 Then
                If .TextMatrix(lngRow, .ColIndex("效期")) <> "" Then
                    If Not IsDate(.TextMatrix(lngRow, .ColIndex("效期"))) Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("效期"), lngRow, .ColIndex("效期")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt效期 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【效期】列"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "格式错误"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【效期】列应为日期格式，如3000-01-01、3000/01/01或者30000101！"
                    End If
                End If
            End If
             
            '灭菌日期
            If GetColumnPostation("灭菌日期") > 0 Then
                If .TextMatrix(lngRow, .ColIndex("灭菌日期")) <> "" Then
                    If Not IsDate(.TextMatrix(lngRow, .ColIndex("灭菌日期"))) Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("灭菌日期"), lngRow, .ColIndex("灭菌日期")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt灭菌日期 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【灭菌日期】列"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "格式错误"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【灭菌日期】列应为日期格式，如3000-01-01、3000/01/01或者30000101！"
                    End If
                End If
            End If
            '灭菌效期
            If GetColumnPostation("灭菌效期") > 0 Then
                If .TextMatrix(lngRow, .ColIndex("灭菌效期")) <> "" Then
                    If Not IsDate(.TextMatrix(lngRow, .ColIndex("灭菌效期"))) Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("灭菌效期"), lngRow, .ColIndex("灭菌效期")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt灭菌效期 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【灭菌效期】列"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "格式错误"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【灭菌效期】列应为日期格式，如3000-01-01、3000/01/01或者30000101！"
                    End If
                End If
            End If
            '发票日期
            If GetColumnPostation("发票日期") > 0 Then
                If .TextMatrix(lngRow, .ColIndex("发票日期")) <> "" Then
                    If Not IsDate(.TextMatrix(lngRow, .ColIndex("发票日期"))) Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("发票日期"), lngRow, .ColIndex("发票日期")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt发票日期 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【发票日期】列"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "格式错误"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【发票日期】列应为是日期格式，如3000-01-01、3000/01/01或者30000101！"
                    End If
                End If
            End If
            If GetColumnPostation("生产日期") > 0 Then
                If .TextMatrix(lngRow, .ColIndex("生产日期")) <> "" Then
                    If Not IsDate(.TextMatrix(lngRow, .ColIndex("生产日期"))) Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("生产日期"), lngRow, .ColIndex("生产日期")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt生产日期 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【生产日期】列"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "格式错误"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "【生产日期】列应为日期格式，如3000-01-01、3000/01/01或者30000101！"
                    End If
                End If
            End If
            '检查存储库房
            If .TextMatrix(lngRow, .ColIndex("卫材编码")) <> "" Then
                gstrSQL = "Select 1" & vbNewLine & _
                            "From 收费项目目录 A, 收费执行科室 B" & vbNewLine & _
                            "Where a.Id = b.收费细目id And b.执行科室id = [1] And a.类别 = '4' And a.编码 = [2] and rownum<2"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "存储库房", mlngStockID, .TextMatrix(lngRow, .ColIndex("卫材编码")))
                If rsTemp.RecordCount = 0 Then
                    .Cell(flexcpForeColor, lngRow, .ColIndex("卫材编码"), lngRow, .ColIndex("卫材编码")) = vbRed
                    vsfError.Rows = vsfError.Rows + 1
                    vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt存储库房 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                    vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【卫材编码】列"
                    vsfError.TextMatrix(vsfError.Rows - 1, 2) = "数据错误"
                    vsfError.TextMatrix(vsfError.Rows - 1, 3) = "该卫材未在" & mstrStock & "库房中设置存储状态，请在卫材目录中调整存储状态！"
                End If
                '高值卫材需要检查其他特性
                If mblnVirtualStock = True Then
                    gstrSQL = "Select b.高值材料, b. 跟踪病人, b.跟踪在用, b.在用分批" & vbNewLine & _
                        "From 收费项目目录 A, 材料特性 B" & vbNewLine & _
                        "Where a.Id = b.材料id And a.类别 = '4' And a.编码 = [1]"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "存储库房", .TextMatrix(lngRow, .ColIndex("卫材编码")))
                        If (zlStr.Nvl(rsTemp!高值材料, 0) = 0 Or zlStr.Nvl(rsTemp!跟踪病人, 0) = 0 Or zlStr.Nvl(rsTemp!跟踪在用, 0) = 0 Or zlStr.Nvl(rsTemp!在用分批, 0) = 0) Then
                            .Cell(flexcpForeColor, lngRow, .ColIndex("卫材编码"), lngRow, .ColIndex("卫材编码")) = vbRed
                            vsfError.Rows = vsfError.Rows + 1
                            vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt虚拟库房 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                            vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【卫材编码】列"
                            vsfError.TextMatrix(vsfError.Rows - 1, 2) = "数据错误"
                            vsfError.TextMatrix(vsfError.Rows - 1, 3) = "虚拟库房卫材必须是高值材料、跟踪在用、跟踪病人、在用分批，请到卫材目录中修改该卫材属性！"
                        End If
                End If
            End If
            '商品条码检查
            If GetColumnPostation("商品码") > 0 Then
                If .TextMatrix(lngRow, .ColIndex("商品码")) <> "" Then
                    gstrSQL = "Select 1" & vbNewLine & _
                                "From 药品收发记录 A, 收费项目目录 B" & vbNewLine & _
                                "Where a.药品id = b.Id And b.编码 = [1] And b.类别 = '4' And a.库房id = [2] And a.商品条码 =[3] And Rownum < 2"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "商品码检查", .TextMatrix(lngRow, .ColIndex("卫材编码")), mlngStockID, .TextMatrix(lngRow, .ColIndex("商品码")))
                    If rsTemp.RecordCount <> 0 Then
                        .Cell(flexcpForeColor, lngRow, .ColIndex("商品码"), lngRow, .ColIndex("商品码")) = vbRed
                        vsfError.Rows = vsfError.Rows + 1
                        vsfError.Cell(flexcpPicture, vsfError.Rows - 1, 0, vsfError.Rows - 1, 0) = IIf(mbyt商品条码 = 0, imgError.ListImages(2).Picture, imgError.ListImages(1).Picture)
                        vsfError.TextMatrix(vsfError.Rows - 1, 1) = lngRow & "行【商品码】列"
                        vsfError.TextMatrix(vsfError.Rows - 1, 2) = "数据错误"
                        vsfError.TextMatrix(vsfError.Rows - 1, 3) = "商品条码重复，请修改！"
                    End If
                End If
            End If
            If ProCheck.Value + 100 / (vsfList.Rows - 1) >= 100 Then
                ProCheck.Value = 100
            Else
                ProCheck.Value = ProCheck.Value + 100 / (vsfList.Rows - 1)
            End If
        Next
    End With
    Set cbrControl = Me.cbsMain(1).Controls.Find(xtpControlButton, MCONTOOLSAVE, , True)
    cbrControl.Enabled = True
    With vsfError
        If .Rows > 1 Then
            If mbyt导入方式 = 0 Then
                cbrControl.Enabled = True
            Else
                For lngRow = 1 To .Rows - 1
                    If .Cell(flexcpPicture, lngRow, 0, lngRow, 0) = imgError.ListImages(1).Picture Or .Cell(flexcpPicture, lngRow, 0, lngRow, 0) = imgError.ListImages(2).Picture Then
                        cbrControl.Enabled = False
                        Exit For
                    End If
                Next
            End If
        End If
    End With
    If vsfList.Rows > 1 Then
        Set cbrControl = Me.cbsMain(1).Controls.Find(xtpControlButton, MCONTOOLOUTPUT, , True)
        cbrControl.Enabled = True
        Set cbrControl = Me.cbsMain(1).Controls.Find(xtpControlButton, MCONTOOLCHECK, , True)
        cbrControl.Enabled = True
    End If
    Exit Sub
    
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetColumn()
    '列宽度对齐方式设置
    Dim lngRow As Long
    Dim lngCol As Long
    
    With vsfList
        For lngCol = 1 To .Cols - 1
            Select Case .TextMatrix(0, lngCol)
                Case "卫材编码", "卫材名称", "规格", "产地", "批号", "商品码", "生产日期", "效期", "灭菌效期", "灭菌日期", "发票日期"
                    .ColAlignment(lngCol) = flexAlignLeftCenter
                Case "数量", "成本价", "成本金额", "发票金额"
                    .ColAlignment(lngCol) = flexAlignRightCenter
                Case Else
                    .ColAlignment(lngCol) = flexAlignRightCenter
            End Select
        Next
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        .ColWidth(0) = 300
        If GetColumnPostation("卫材编码") > 0 Then
            .ColWidth(.ColIndex("卫材编码")) = 2000
        End If
        If GetColumnPostation("生产日期") > 0 Then
            .ColWidth(.ColIndex("生产日期")) = 1000
        End If
        If GetColumnPostation("效期") > 0 Then
            .ColWidth(.ColIndex("效期")) = 1000
        End If
        If GetColumnPostation("灭菌效期") > 0 Then
            .ColWidth(.ColIndex("灭菌效期")) = 1000
        End If
        If GetColumnPostation("灭菌日期") > 0 Then
            .ColWidth(.ColIndex("灭菌日期")) = 1000
        End If
        If GetColumnPostation("发票日期") > 0 Then
            .ColWidth(.ColIndex("发票日期")) = 1000
        End If
    End With
End Sub

Private Sub InitControlPosition()
    '控件位置控制
    On Error Resume Next
    
    lblFile.Move 100, 600
    txtFile.Move lblFile.Left + lblFile.Width + 20, lblFile.Top - 40, Me.ScaleWidth - (cmdFile.Width + txtFile.Left)
    cmdFile.Move txtFile.Left + txtFile.Width, txtFile.Top - 30
    lblProvider.Move lblFile.Left, lblFile.Top + lblFile.Height + 200
    txtProvider.Move txtFile.Left, lblProvider.Top - 50, txtFile.Width
    
    If mlngModule = 1712 Then
        cboIOType.Move txtFile.Left, lblProvider.Top, txtFile.Width
        cmdProvider.Move txtProvider.Left + txtProvider.Width - 10, cboIOType.Top - 60
    Else
        cboIOType.Move txtFile.Left, lblProvider.Top, txtFile.Width + cmdFile.Width
        cmdProvider.Visible = False
    End If
    vsfList.Move lblFile.Left, txtProvider.Top + txtProvider.Height + 50, Me.Width - lblFile.Left - vsfList.Left - 120, ((Me.ScaleHeight - vsfList.Top) / 4) * 3
    picSplit.Move lblFile.Left, vsfList.Top + vsfList.Height + 50, Me.ScaleWidth - lblFile.Left - vsfList.Left
    lblCollect.Width = picSplit.Width
    
    vsfError.Move lblFile.Left, picSplit.Top + picSplit.Height + 50, Me.ScaleWidth - lblFile.Left, Me.ScaleHeight - picSplit.Top - picSplit.Height - 100
    ProCheck.Move lblFile.Left, (vsfError.Top + vsfError.Height) / 2, Me.ScaleWidth - lblFile.Left - ProCheck.Left
    ProCheck.Visible = False
End Sub

Private Sub Form_Resize()
    Call InitControlPosition
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mobjWS = Nothing
    Set mobjWB = Nothing
    Set mobjXLS = Nothing
End Sub

Private Sub picSplit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    With picSplit
        If .Top + y < 2000 Then Exit Sub
        If .Top + y > ScaleHeight - 2000 Then Exit Sub
        .Move .Left, .Top + y
    End With

    With vsfList
        .Height = picSplit.Top - .Top
    End With
    
    With vsfError
        .Top = picSplit.Top + picSplit.Height + 100
        .Height = ScaleHeight - .Top
    End With
    Me.Refresh
End Sub

Private Sub txtFile_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If Trim(txtFile.Text) <> "" Then OS.PressKey vbKeyTab
End Sub

Private Sub txtFile_LostFocus()
'    If Trim(txtFile.Text) <> "" Then OS.PressKey vbKeyTab
End Sub

Private Sub txtProvider_Change()
    txtProvider.Tag = ""
End Sub

Private Sub txtProvider_GotFocus()
    OS.OpenIme False
    zlControl.TxtSelAll txtProvider
End Sub

Private Sub txtProvider_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Val(txtProvider.Tag) <> 0 Then OS.PressKey vbKeyTab: Exit Sub
    If Val(txtProvider.Tag) = 0 And Trim(txtProvider.Text) = "" Then Exit Sub
    If Select供应商(Me, txtProvider, Trim(txtProvider.Text)) = False Then Exit Sub
    OS.PressKey vbKeyTab
End Sub

Private Sub txtProvider_LostFocus()
    If txtProvider.Text = "" Then
        txtProvider.Tag = "0"
        Exit Sub
    End If
End Sub

Private Sub ImportData()
'将外部文件的数据导入到药品收发记录表中
    Dim lngCol As Long, lngRow As Long
    Dim lngMaterialID As Long
    Dim dblTransVal As Double, dblAddRate As Double
    Dim dblSalePrice As Double, dblSale As Double, dblCostPrice As Double, dblCost As Double, dblCurSale As Double
    Dim dblQTY As Double
    Dim arrCols As Variant
    Dim strTmp As String, strInsert As String, strNo As String, strMess As String
    Dim rsTmp As ADODB.Recordset
    Dim strPlaceProduction As String, strPackageUnit As String
    Dim bytLotPrice As Byte
    Dim blnLot As Boolean, blnOnce As Boolean
    Dim blnTran As Boolean  '记录事物是否开始了
    
    On Error GoTo ErrHandle
    
    With vsfList
        '单据号
        Select Case mlngModule
            Case 1712
                strNo = sys.GetNextNo(68, mlngStockID)
            Case 1714
                strNo = sys.GetNextNo(70, mlngStockID)
        End Select
        '组织数据
        
        On Error GoTo ErrHandle
        blnTran = True
        gcnOracle.BeginTrans
        For lngRow = 1 To .Rows - 1
            '检查卫材编码
            gstrSQL = "select a.ID, a.是否变价, 1/(1-b.指导差价率/100)-1 加成率, b.换算系数, b.库房分批" & _
                      ", b.一次性材料, b.高值材料, b. 跟踪病人, b.跟踪在用, b.在用分批, b.包装单位, c.现价 " & _
                      "from 收费项目目录 a, 材料特性 b, 收费价目 c " & _
                      "where a.ID=b.材料ID and a.ID=c.收费细目ID and a.编码=[1] and a.类别='4' " & _
                      " and a.撤档时间>=to_date('3000-1-1','yyyy-mm-dd') and c.终止日期=to_date('3000-1-1','yyyy-mm-dd') " & _
                      GetPriceClassString("C")
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, "" & .TextMatrix(lngRow, .ColIndex("卫材编码")))
            '虚拟库房的库房必须具有[高值材料][跟踪病人][跟踪在用][在用分批]属性
            If rsTmp.RecordCount > 0 Then
                lngMaterialID = rsTmp!Id
                dblCurSale = IIf(IsNull(rsTmp!现价), 0, rsTmp!现价)
                dblTransVal = IIf(IsNull(rsTmp!换算系数), 1, rsTmp!换算系数)
                dblAddRate = IIf(IsNull(rsTmp!加成率), 0, rsTmp!加成率)
                bytLotPrice = IIf(IsNull(rsTmp!是否变价), 0, rsTmp!是否变价)
                blnLot = IIf(IsNull(rsTmp!库房分批), 0, rsTmp!库房分批)
                blnOnce = IIf(IsNull(rsTmp!一次性材料), 0, rsTmp!一次性材料)
                strPackageUnit = IIf(IsNull(rsTmp!包装单位), "", rsTmp!包装单位)
                
                If mlngModule = 1712 Then
                    strInsert = "zl_材料外购_INSERT("
                    'NO
                    strInsert = strInsert & "'" & strNo & "',"
                    '序号
                    strInsert = strInsert & lngRow - 1 & ","
                    '库房ID
                    strInsert = strInsert & mlngStockID & ","
                    '供应商ID
                    strInsert = strInsert & txtProvider.Tag & ","
                    '材料ID
                    strInsert = strInsert & lngMaterialID & ","
                    '产地
                    If GetColumnPostation("产地") > 0 Then
                        strTmp = UCase(Trim(.TextMatrix(lngRow, .ColIndex("产地"))))
                    Else
                        strTmp = ""
                    End If
                    strInsert = strInsert & IIf(strTmp = "", "null", "'" & strTmp & "'") & ","
                    '批号
                    If GetColumnPostation("批号") > 0 Then
                        strTmp = Trim(.TextMatrix(lngRow, .ColIndex("批号")))
                    Else
                        strTmp = ""
                    End If
                    strInsert = strInsert & IIf(strTmp = "", "null", "'" & strTmp & "'") & ","
                    '生产日期
                    If GetColumnPostation("生产日期") > 0 Then
                        strTmp = IIf(IsDate(Trim(.TextMatrix(lngRow, .ColIndex("生产日期")))) = True, Trim(.TextMatrix(lngRow, .ColIndex("生产日期"))), "")
                    Else
                        strTmp = ""
                    End If
                    If IsNumeric(strTmp) Then
                        strTmp = TranNumToDate(strTmp)
                    End If
                    strInsert = strInsert & IIf(strTmp = "", "null", "to_date('" & Format(strTmp, "yyyy-mm-dd") & "','YYYY-MM-DD')") & ","
                    '效期
                    If blnLot Then  '库房分批才有效期
                        If GetColumnPostation("效期") > 0 Then
                            strTmp = IIf(IsDate(Trim(.TextMatrix(lngRow, .ColIndex("效期")))) = True, Trim(.TextMatrix(lngRow, .ColIndex("效期"))), "")
                        Else
                            strTmp = ""
                        End If
                        If IsNumeric(strTmp) Then
                            strTmp = TranNumToDate(strTmp)
                        End If
                        strInsert = strInsert & IIf(strTmp = "", "null", "to_date('" & Format(strTmp, "yyyy-mm-dd") & "','YYYY-MM-DD')") & ","
                    Else
                        strInsert = strInsert & "null,"
                    End If
                    If blnOnce Then
                        '灭菌日期
                        If GetColumnPostation("灭菌日期") > 0 Then
                            strTmp = IIf(IsDate(Trim(.TextMatrix(lngRow, .ColIndex("灭菌日期")))) = True, Trim(.TextMatrix(lngRow, .ColIndex("灭菌日期"))), "")
                        Else
                            strTmp = ""
                        End If
                        If IsNumeric(strTmp) Then
                            strTmp = TranNumToDate(strTmp)
                        End If
                        strInsert = strInsert & IIf(strTmp = "", "null", "to_date('" & Format(strTmp, "yyyy-mm-dd") & "','YYYY-MM-DD')") & ","
                        '灭菌效期
                        If GetColumnPostation("灭菌效期") > 0 Then
                            strTmp = IIf(IsDate(Trim(.TextMatrix(lngRow, .ColIndex("灭菌效期")))) = True, Trim(.TextMatrix(lngRow, .ColIndex("灭菌效期"))), "")
                        Else
                            strTmp = ""
                        End If
                        If IsNumeric(strTmp) Then
                            strTmp = TranNumToDate(strTmp)
                        End If
                        strInsert = strInsert & IIf(strTmp = "", "null", "to_date('" & Format(strTmp, "yyyy-mm-dd") & "','YYYY-MM-DD')") & ","
                    Else
                        strInsert = strInsert & "null,null,"
                    End If
                    '数量
                    If GetColumnPostation("数量") > 0 Then
                        strTmp = IIf(IsNumeric(Trim(.TextMatrix(lngRow, .ColIndex("数量")))) = True, Trim(.TextMatrix(lngRow, .ColIndex("数量"))), "")
                    Else
                        strTmp = ""
                    End If
                    dblQTY = IIf(strTmp = "", 0, strTmp)
                    strInsert = strInsert & GetFormat(dblQTY * dblTransVal, g_小数位数.obj_最大小数.数量小数) & ","
                    '成本价
                    If GetColumnPostation("成本价") > 0 Then
                        strTmp = IIf(IsNumeric(Trim(.TextMatrix(lngRow, .ColIndex("成本价")))) = True, Trim(.TextMatrix(lngRow, .ColIndex("成本价"))), "")
                    Else
                        strTmp = ""
                    End If
                    dblCostPrice = CDbl(IIf(strTmp = "", 0, strTmp))
                    strInsert = strInsert & GetFormat(dblCostPrice / IIf(dblTransVal = 0, 1, dblTransVal), g_小数位数.obj_最大小数.成本价小数) & ","
                    '成本金额
                    If GetColumnPostation("成本金额") > 0 Then
                        strTmp = IIf(IsNumeric(Trim(.TextMatrix(lngRow, .ColIndex("成本金额")))) = True, Trim(.TextMatrix(lngRow, .ColIndex("成本金额"))), "")
                    Else
                        strTmp = ""
                    End If
                    dblCost = CDbl(IIf(strTmp = "", 0, strTmp))
                    strInsert = strInsert & GetFormat(dblCost, g_小数位数.obj_最大小数.金额小数) & ","
                    '扣率
                    strInsert = strInsert & "100,"
                    '售价
                    dblSalePrice = IIf(bytLotPrice = 1, dblCostPrice / IIf(dblTransVal = 0, 1, dblTransVal) * (dblAddRate + 1), dblCurSale)
                    strInsert = strInsert & GetFormat(dblSalePrice, g_小数位数.obj_最大小数.零售价小数) & ","
                    '售价金额
                    dblSale = GetFormat(dblQTY * dblTransVal, g_小数位数.obj_最大小数.数量小数) * GetFormat(dblSalePrice, g_小数位数.obj_最大小数.零售价小数)
                    strInsert = strInsert & GetFormat(dblSale, g_小数位数.obj_最大小数.金额小数) & ","
                    '差价
                    strInsert = strInsert & GetFormat(dblSale, g_小数位数.obj_最大小数.金额小数) - GetFormat(dblCost, g_小数位数.obj_最大小数.金额小数) & ","
                    '零售差价，摘要，注册证号
                    strInsert = strInsert & "null,null,null,"
                    '填制人
                    strInsert = strInsert & "'" & gstrUserName & "',"
                    '随货单号
                    strInsert = strInsert & "null,"
                    '发票号
                    If GetColumnPostation("发票号") > 0 Then
                        strTmp = Trim(.TextMatrix(lngRow, .ColIndex("发票号")))
                    Else
                        strTmp = ""
                    End If
                    strInsert = strInsert & IIf(strTmp = "", "null", "'" & strTmp & "'") & ","
                    '发票日期
                    If GetColumnPostation("发票日期") > 0 Then
                        strTmp = IIf(IsDate(Trim(.TextMatrix(lngRow, .ColIndex("发票日期")))) = True, Trim(.TextMatrix(lngRow, .ColIndex("发票日期"))), "")
                    Else
                        strTmp = ""
                    End If
                    If IsNumeric(strTmp) Then
                        strTmp = TranNumToDate(strTmp)
                    End If
                    strInsert = strInsert & IIf(strTmp = "", "null", "to_date('" & Format(strTmp, "yyyy-mm-dd") & "','YYYY-MM-DD')") & ","
                    '发票金额
                    If GetColumnPostation("发票金额") > 0 Then
                        strTmp = IIf(IsNumeric(Trim(.TextMatrix(lngRow, .ColIndex("发票金额")))) = True, Trim(.TextMatrix(lngRow, .ColIndex("发票金额"))), "")
                    Else
                        strTmp = ""
                    End If
                    strInsert = strInsert & IIf(strTmp = "", 0, strTmp) & ","
                    '填制日期
                    strInsert = strInsert & "to_date('" & Format(Now(), "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd HH24:MI:SS'),"
                    '核查人，核查日期，批次，退货，高值材料
                    strInsert = strInsert & "null,null,null,1,null,"
                    '商品码
                    If GetColumnPostation("商品码") > 0 Then
                        strTmp = Trim(.TextMatrix(lngRow, .ColIndex("商品码")))
                    Else
                        strTmp = ""
                    End If
                    strInsert = strInsert & IIf(strTmp = "", "null)", "'" & strTmp & "')")
                ElseIf mlngModule = 1714 Then
                    strInsert = "zl_材料其他入库_INSERT("
                    'No
                    strInsert = strInsert & "'" & strNo & "',"
                    '序号
                    strInsert = strInsert & lngRow - 1 & ","
                    '库房id
                    strInsert = strInsert & mlngStockID & ","
                    '入出类别
                    strInsert = strInsert & cboIOType.ItemData(cboIOType.ListIndex) & ","
                    '材料id
                    strInsert = strInsert & lngMaterialID & ","
                    '数量
                    If GetColumnPostation("数量") > 0 Then
                        strTmp = IIf(IsNumeric(Trim(.TextMatrix(lngRow, .ColIndex("数量")))) = True, Trim(.TextMatrix(lngRow, .ColIndex("数量"))), "")
                    Else
                        strTmp = ""
                    End If
                    dblQTY = IIf(strTmp = "", 0, strTmp)
                    strInsert = strInsert & GetFormat(dblQTY * dblTransVal, g_小数位数.obj_最大小数.数量小数) & ","
                    '成本价
                    If GetColumnPostation("成本价") > 0 Then
                        strTmp = IIf(IsNumeric(Trim(.TextMatrix(lngRow, .ColIndex("成本价")))) = True, Trim(.TextMatrix(lngRow, .ColIndex("成本价"))), "")
                    Else
                        strTmp = ""
                    End If
                    dblCostPrice = CDbl(IIf(strTmp = "", 0, strTmp))
                    strInsert = strInsert & GetFormat(dblCostPrice / IIf(dblTransVal = 0, 1, dblTransVal), g_小数位数.obj_最大小数.成本价小数) & ","
                    '成本金额
                    If GetColumnPostation("成本金额") > 0 Then
                        strTmp = IIf(IsNumeric(Trim(.TextMatrix(lngRow, .ColIndex("成本金额")))) = True, Trim(.TextMatrix(lngRow, .ColIndex("成本金额"))), "")
                    Else
                        strTmp = ""
                    End If
                    dblCost = CDbl(IIf(strTmp = "", 0, strTmp))
                    strInsert = strInsert & GetFormat(dblCost, g_小数位数.obj_最大小数.金额小数) & ","
                    '售价
                    dblSalePrice = IIf(bytLotPrice = 1, dblCostPrice / IIf(dblTransVal = 0, 1, dblTransVal) * (dblAddRate + 1), dblCurSale)
                    strInsert = strInsert & GetFormat(dblSalePrice, g_小数位数.obj_最大小数.零售价小数) & ","
                    '售价金额
                    dblSale = GetFormat(dblQTY * dblTransVal, g_小数位数.obj_最大小数.数量小数) * GetFormat(dblSalePrice, g_小数位数.obj_最大小数.零售价小数)
                    strInsert = strInsert & GetFormat(dblSale, g_小数位数.obj_最大小数.金额小数) & ","
                    '差价
                    strInsert = strInsert & GetFormat(dblSale, g_小数位数.obj_最大小数.金额小数) - GetFormat(dblCost, g_小数位数.obj_最大小数.金额小数) & ","
                    '零售差价
                    strInsert = strInsert & "null,"
                    '填制人
                    strInsert = strInsert & "'" & gstrUserName & "',"
                    '填制日期
                    strInsert = strInsert & "to_date('" & Format(Now(), "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd HH24:MI:SS'),"
                    '摘要
                    strInsert = strInsert & "null,"
                    '产地
                    If GetColumnPostation("产地") > 0 Then
                        strTmp = UCase(Trim(.TextMatrix(lngRow, .ColIndex("产地"))))
                    Else
                        strTmp = ""
                    End If
                    strInsert = strInsert & IIf(strTmp = "", "null", "'" & strTmp & "'") & ","
                    '批号
                    If GetColumnPostation("批号") > 0 Then
                        strTmp = Trim(.TextMatrix(lngRow, .ColIndex("批号")))
                    Else
                        strTmp = ""
                    End If
                    strInsert = strInsert & IIf(strTmp = "", "null", "'" & strTmp & "'") & ","
                    '生产日期
                    If GetColumnPostation("生产日期") > 0 Then
                        strTmp = IIf(IsDate(Trim(.TextMatrix(lngRow, .ColIndex("生产日期")))) = True, Trim(.TextMatrix(lngRow, .ColIndex("生产日期"))), "")
                    Else
                        strTmp = ""
                    End If
                    If IsNumeric(strTmp) Then
                        strTmp = TranNumToDate(strTmp)
                    End If
                    strInsert = strInsert & IIf(strTmp = "", "null", "to_date('" & Format(strTmp, "yyyy-mm-dd") & "','YYYY-MM-DD')") & ","
                    '效期
                    If blnLot Then  '库房分批才有效期
                        If GetColumnPostation("效期") > 0 Then
                             strTmp = IIf(IsDate(Trim(.TextMatrix(lngRow, .ColIndex("效期")))) = True, Trim(.TextMatrix(lngRow, .ColIndex("效期"))), "")
                        Else
                            strTmp = ""
                        End If
                        If IsNumeric(strTmp) Then
                            strTmp = TranNumToDate(strTmp)
                        End If
                        strInsert = strInsert & IIf(strTmp = "", "null", "to_date('" & Format(strTmp, "yyyy-mm-dd") & "','YYYY-MM-DD')") & ","
                    Else
                        strInsert = strInsert & "null,"
                    End If
                    If blnOnce Then
                        '灭菌日期
                        If GetColumnPostation("灭菌日期") > 0 Then
                            strTmp = IIf(IsDate(Trim(.TextMatrix(lngRow, .ColIndex("灭菌日期")))) = True, Trim(.TextMatrix(lngRow, .ColIndex("灭菌日期"))), "")
                        Else
                            strTmp = ""
                        End If
                        If IsNumeric(strTmp) Then
                            strTmp = TranNumToDate(strTmp)
                        End If
                        strInsert = strInsert & IIf(strTmp = "", "null", "to_date('" & Format(strTmp, "yyyy-mm-dd") & "','YYYY-MM-DD')") & ","
                        '灭菌效期
                        If GetColumnPostation("灭菌效期") > 0 Then
                            strTmp = IIf(IsDate(Trim(.TextMatrix(lngRow, .ColIndex("灭菌效期")))) = True, Trim(.TextMatrix(lngRow, .ColIndex("灭菌效期"))), "")
                        Else
                            strTmp = ""
                        End If
                        strInsert = strInsert & IIf(strTmp = "", "null )", "to_date('" & Format(strTmp, "yyyy-mm-dd") & "','YYYY-MM-DD') )")
                    Else
                        strInsert = strInsert & "null,null)"
                    End If
                End If
                Call zlDatabase.ExecuteProcedure(strInsert, Me.Caption)
            End If
        Next
    End With
    
    If blnTran = True Then
        gcnOracle.CommitTrans
    End If
    mblnResult = True
    MsgBox "保存成功！", vbInformation, gstrSysName
    Exit Sub
    
ErrHandle:
    If blnTran = True Then
        gcnOracle.RollbackTrans
    End If
    If ErrCenter() = 1 Then
        Resume
    End If
    MsgBox "保存成功！", vbInformation, gstrSysName
    Call SaveErrLog
End Sub

Private Function GetColIndex(ByVal strColName As String) As Long
    Dim i As Long
    For i = 1 To mobjWS.UsedRange.Columns.count
        If mobjWS.UsedRange.cells(1, i) = strColName Then
            GetColIndex = i
            Exit Function
        End If
    Next
End Function

Private Sub vsfError_EnterCell()
    Dim strTemp As String
    Dim lngRow As Long
    Dim lngCol As Long
    Dim strCol As String
    
    With vsfError
        If .Row = 0 Then Exit Sub
        .FocusRect = flexFocusSolid
        If InStr(1, .TextMatrix(.Row, 1), "列") = 0 Then Exit Sub
        If .TextMatrix(.Row, 1) <> "" Then
            strTemp = .TextMatrix(.Row, 1)
            lngRow = Mid(strTemp, 1, InStr(1, strTemp, "行") - 1)
            strCol = Mid(strTemp, InStr(1, strTemp, "【") + 1, InStr(1, strTemp, "】") - InStr(1, strTemp, "【") - 1)
            lngCol = vsfList.ColIndex(strCol)
            If lngRow > vsfList.Rows - 1 Then MsgBox "改行数据已经被删除了！", vbInformation, gstrSysName: Exit Sub
            vsfList.Row = lngRow
            vsfList.Col = lngCol
            vsfList.ShowCell lngRow, lngCol
        End If
    End With
End Sub

Private Sub vsfList_DblClick()
    With vsfList
        .EditCell
        .EditSelStart = 0
        .EditSelLength = Len(.EditText)
    End With
End Sub

Private Sub vsfList_EnterCell()
    With vsfList
        If .Row < 1 Then Exit Sub
        .FocusRect = flexFocusSolid
    End With
End Sub

Private Sub vsfList_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim bln成本金额 As Boolean
    Dim bln发票金额 As Boolean
    Dim dbl成本金额 As Double
    Dim dbl发票金额 As Double
    Dim lngRow As Long
    
    If KeyCode = vbKeyDelete Then
        If MsgBox("将删除第" & vsfList.Row & "行数据是否继续？", vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            With vsfList
                .RemoveItem .Row
                bln成本金额 = IIf(GetColumnPostation("成本金额") > 0, True, False)
                bln发票金额 = IIf(GetColumnPostation("发票金额") > 0, True, False)
                For lngRow = 1 To .Rows - 1
                    If bln成本金额 = True Then
                        dbl成本金额 = dbl成本金额 + Val(.TextMatrix(lngRow, .ColIndex("成本金额")))
                    End If
                    If bln发票金额 = True Then
                        dbl发票金额 = dbl发票金额 + Val(.TextMatrix(lngRow, .ColIndex("发票金额")))
                    End If
                Next
                lblCollect.Caption = "成本金额：" & Format(dbl成本金额, mFMT.FM_金额) & "元          发票金额：" & Format(dbl发票金额, mFMT.FM_金额) & "元"
            End With
        End If
    End If
End Sub

Private Sub vsfList_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim lngRow As Long
    Dim lngCol As Long
    Dim dbl成本金额 As Double
    Dim dbl发票金额 As Double
    Dim bln成本金额 As Boolean
    Dim bln发票金额 As Boolean
    Dim strTemp As String
    
    Dim cbrControl As CommandBarControl
    If mbyt导入方式 = 1 Then
        Set cbrControl = Me.cbsMain(1).Controls.Find(xtpControlButton, MCONTOOLSAVE, , True)
        cbrControl.Enabled = False
    End If
    
    With vsfList
        strTemp = .EditText
        If .TextMatrix(0, Col) = "成本金额" Then
            strTemp = Format(Val(strTemp), mFMT.FM_金额)
            dbl成本金额 = 0
            dbl发票金额 = 0
            bln发票金额 = IIf(GetColumnPostation("发票金额") > 0, True, False)
            For lngRow = 1 To .Rows - 1
                If lngRow = Row Then
                    dbl成本金额 = dbl成本金额 + strTemp
                Else
                    dbl成本金额 = dbl成本金额 + Val(.TextMatrix(lngRow, .ColIndex("成本金额")))
                End If
                If bln发票金额 = True Then
                    dbl发票金额 = dbl发票金额 + Val(.TextMatrix(lngRow, .ColIndex("发票金额")))
                End If
            Next
            .EditText = strTemp
        End If
        If .TextMatrix(0, Col) = "发票金额" Then
            strTemp = Format(Val(strTemp), mFMT.FM_金额)
            dbl成本金额 = 0
            dbl发票金额 = 0
            bln成本金额 = IIf(GetColumnPostation("成本金额") > 0, True, False)
            For lngRow = 1 To .Rows - 1
                If lngRow = Row Then
                    dbl发票金额 = dbl发票金额 + strTemp
                Else
                    dbl发票金额 = dbl发票金额 + Val(.TextMatrix(lngRow, .ColIndex("发票金额")))
                End If
                If bln成本金额 = True Then
                    dbl成本金额 = dbl成本金额 + Val(.TextMatrix(lngRow, .ColIndex("成本金额")))
                End If
            Next
            .EditText = strTemp
        End If
        If .TextMatrix(0, Col) = "发票金额" Or .TextMatrix(0, Col) = "成本金额" Then
            lblCollect.Caption = "成本金额：" & Format(dbl成本金额, mFMT.FM_金额) & "元          发票金额：" & Format(dbl发票金额, mFMT.FM_金额) & "元"
        End If
        If .TextMatrix(0, Col) = "数量" Then
            .EditText = Format(Val(strTemp), mFMT.FM_数量)
        End If
        If .TextMatrix(0, Col) = "成本价" Then
            .EditText = Format(Val(strTemp), mFMT.FM_成本价)
        End If
        If .TextMatrix(0, Col) = "生产日期" Then
            .EditText = FormatDate(strTemp)
        End If
        If .TextMatrix(0, Col) = "效期" Then
            .EditText = FormatDate(strTemp)
        End If
        If .TextMatrix(0, Col) = "灭菌日期" Then
            .EditText = FormatDate(strTemp)
        End If
        If .TextMatrix(0, Col) = "灭菌效期" Then
            .EditText = FormatDate(strTemp)
        End If
        If .TextMatrix(0, Col) = "发票日期" Then
            .EditText = FormatDate(strTemp)
        End If
        
    End With
End Sub

