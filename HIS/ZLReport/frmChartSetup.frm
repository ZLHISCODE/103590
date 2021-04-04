VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0BE3824E-5AFE-4B11-A6BC-4B3AD564982A}#8.0#0"; "olch2x8.ocx"
Begin VB.Form frmChartSetup 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "图表设置"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6885
   Icon            =   "frmChartSetup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CheckBox chkFormat 
      Caption         =   "XY轴交换"
      Height          =   195
      Index           =   1
      Left            =   2186
      TabIndex        =   25
      Top             =   4245
      Width           =   1020
   End
   Begin VB.CheckBox chkFormat 
      Caption         =   "三维效果"
      Height          =   195
      Index           =   0
      Left            =   990
      TabIndex        =   24
      Top             =   4245
      Width           =   1020
   End
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   3060
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3390
      Width           =   330
   End
   Begin VB.CommandButton cmdFore 
      BackColor       =   &H00000000&
      Height          =   300
      Left            =   1770
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3390
      Width           =   330
   End
   Begin VB.CommandButton cmdFont 
      Height          =   315
      Left            =   3060
      Picture         =   "frmChartSetup.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "字体设置"
      Top             =   3000
      Width           =   330
   End
   Begin VB.TextBox txtFont 
      Height          =   300
      Left            =   990
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   3015
      Width           =   2085
   End
   Begin VB.TextBox txtFontTitle 
      Enabled         =   0   'False
      Height          =   300
      Left            =   990
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   2625
      Width           =   2085
   End
   Begin MSComDlg.CommonDialog cdg 
      Left            =   240
      Top             =   4695
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   90
      Left            =   -105
      TabIndex        =   34
      Top             =   4530
      Width           =   7605
   End
   Begin VB.ComboBox cboLocate 
      Enabled         =   0   'False
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Top             =   3765
      Width           =   2070
   End
   Begin C1Chart2D8.Chart2D Chart 
      Height          =   3360
      Left            =   3525
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   570
      Width           =   3240
      _Version        =   524288
      _Revision       =   7
      _ExtentX        =   5715
      _ExtentY        =   5927
      _StockProps     =   0
      ControlProperties=   "frmChartSetup.frx":0B14
   End
   Begin VB.CheckBox chkNode 
      Caption         =   "显示结点"
      Height          =   195
      Left            =   5775
      TabIndex        =   28
      Top             =   4245
      Value           =   1  'Checked
      Width           =   1020
   End
   Begin VB.CheckBox chkLine 
      Caption         =   "显示连线"
      Height          =   195
      Left            =   4578
      TabIndex        =   27
      Top             =   4245
      Value           =   1  'Checked
      Width           =   1020
   End
   Begin VB.CheckBox chkSample 
      Alignment       =   1  'Right Justify
      Caption         =   "显示图例"
      Height          =   195
      Left            =   135
      TabIndex        =   22
      Top             =   3825
      Width           =   1050
   End
   Begin VB.CommandButton cmdFontTitle 
      Enabled         =   0   'False
      Height          =   315
      Left            =   3060
      Picture         =   "frmChartSetup.frx":1173
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "字体设置"
      Top             =   2625
      Width           =   330
   End
   Begin VB.TextBox txtTitle 
      Height          =   300
      Left            =   990
      MaxLength       =   50
      TabIndex        =   11
      Top             =   2250
      Width           =   2400
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "应用(&A)"
      Enabled         =   0   'False
      Height          =   350
      Left            =   5430
      TabIndex        =   33
      Top             =   4770
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4140
      TabIndex        =   32
      Top             =   4770
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3015
      TabIndex        =   31
      Top             =   4770
      Width           =   1100
   End
   Begin VB.ComboBox cboStyle 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   990
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1815
      Width           =   2400
   End
   Begin VB.ComboBox cboFY 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   990
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1200
      Width           =   2400
   End
   Begin VB.ComboBox cboFS 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   990
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   840
      Width           =   2400
   End
   Begin VB.ComboBox cboFX 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   990
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   465
      Width           =   2400
   End
   Begin VB.ComboBox cboData 
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   990
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   105
      Width           =   2400
   End
   Begin VB.CheckBox chkGrid 
      Caption         =   "显示网格"
      Height          =   195
      Left            =   3382
      TabIndex        =   26
      Top             =   4245
      Width           =   1020
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "背景颜色"
      Height          =   180
      Left            =   2235
      TabIndex        =   20
      Top             =   3450
      Width           =   720
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "前景颜色"
      Height          =   180
      Left            =   975
      TabIndex        =   18
      Top             =   3450
      Width           =   720
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "图表字体"
      Height          =   180
      Left            =   165
      TabIndex        =   15
      Top             =   3045
      Width           =   720
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "标题字体"
      Height          =   180
      Left            =   165
      TabIndex        =   12
      Top             =   2670
      Width           =   720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   1
      X1              =   90
      X2              =   3390
      Y1              =   1635
      Y2              =   1635
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   0
      X1              =   90
      X2              =   3390
      Y1              =   1650
      Y2              =   1650
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "图表示例："
      Height          =   180
      Left            =   3525
      TabIndex        =   29
      Top             =   255
      Width           =   900
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "标题文本"
      Height          =   180
      Left            =   165
      TabIndex        =   10
      Top             =   2310
      Width           =   720
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "图表样式"
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   165
      TabIndex        =   8
      Top             =   1875
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ｙ值字段"
      Height          =   180
      Left            =   165
      TabIndex        =   6
      Top             =   1260
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "序列字段"
      Height          =   180
      Left            =   165
      TabIndex        =   4
      Top             =   900
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ｘ值字段"
      Height          =   180
      Left            =   165
      TabIndex        =   2
      Top             =   525
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "数据来源"
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   165
      TabIndex        =   0
      Top             =   165
      Width           =   720
   End
End
Attribute VB_Name = "frmChartSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mblnOK As Boolean
Private mobjChart As Object 'byRef:In/Out
Private mobjDatas As RPTDatas 'In
Private mobjItem As RPTItem 'byRef:In/Out
Private mtmpItem As RPTItem

Private Property Let ItemChange(ByVal vData As Boolean)
    cmdApply.Enabled = vData
    If vData Then
        Call SetChartStyleAndData(Chart, mtmpItem)
    End If
End Property

Private Property Get ItemChange() As Boolean
    ItemChange = cmdApply.Enabled
End Property

Public Function ShowMe(frmParent As Object, ByVal objDatas As RPTDatas, objChart As Object, objItem As RPTItem) As Boolean
    Set mobjDatas = objDatas
    Set mobjChart = objChart
    Set mobjItem = objItem
    
    Me.Show 1, frmParent
    If mblnOK Then '窗体卸载或函数完成时,函数中定义变量的引用关系已中断
        Call CopyItem(objItem, mobjItem)
    End If
    ShowMe = mblnOK
End Function

Private Sub cmdApply_Click()
    If Not CheckInput Then Exit Sub
    Call CopyItem(mobjItem, mtmpItem)
    Call SetChartStyleAndData(mobjChart, mobjItem, , , True)
    mblnOK = True
    ItemChange = False
End Sub

Private Sub cmdOK_Click()
    If Not CheckInput Then Exit Sub
    Call CopyItem(mobjItem, mtmpItem)
    Call SetChartStyleAndData(mobjChart, mobjItem, , , True)
    mblnOK = True
    Unload Me
End Sub

Private Sub SetOptionEnabled()
    '0-Plot(散点图),1-Plot(折线图),2-Bar(条形图),3-Pie(饼图),4-StackingBar(层叠图),5-Area(面积图)
    '6-HiLo(股价图-盘高,盘低),7-HiLoOpenClose(股价图-盘高,盘低,开盘,收盘),8-Candle(股价图-阴阳烛图:盘高,盘低,开盘,收盘)
    '9-Polar(级线图),10-Radar(雷达图),11-FilledRadar(填充雷达图),12-Bubble(气泡图)
    
    '仅部份图样有三维样式
    chkFormat(0).Enabled = InStr(",1,2,3,4,5,", "," & cboStyle.ListIndex & ",") > 0
    If Not chkFormat(0).Enabled Then
        chkFormat(0).Value = 0
    End If
    
    '仅部份图样XY轴交换有效
    chkFormat(1).Enabled = InStr(",3,9,10,11,", "," & cboStyle.ListIndex & ",") = 0
    If Not chkFormat(1).Enabled Then
        chkFormat(1).Value = 0
    End If
    
    '饼图无网格
    chkGrid.Enabled = cboStyle.ListIndex <> 3
    If Not chkGrid.Enabled Then chkGrid.Value = 0
    
    '仅部份图样有连线
    chkLine.Enabled = InStr(",2,3,4,5,", "," & cboStyle.ListIndex & ",") = 0
    If Not chkLine.Enabled Then chkLine.Value = 0
    
    '仅部份图样有结点
    chkNode.Enabled = InStr(",2,3,4,5,6,7,8,11,", "," & cboStyle.ListIndex & ",") = 0
    If Not chkNode.Enabled Then chkNode.Value = 0
End Sub

Private Sub chkFormat_Click(Index As Integer)
    Dim i As Integer
    If Visible Then
        mtmpItem.格式 = ""
        For i = 0 To chkFormat.UBound
            mtmpItem.格式 = mtmpItem.格式 & CStr(chkFormat(i).Value)
        Next
        ItemChange = True
    End If
End Sub

Private Sub cboData_Click()
    Dim arrField As Variant, strField As String
    Dim strFX As String, strFY As String, strFS As String
    Dim i As Long
    
    If cboData.ListIndex = -1 Then
        Call CboSetIndex(cboFX.hWnd, -1)
        Call CboSetIndex(cboFS.hWnd, -1)
        Call CboSetIndex(cboFY.hWnd, -1)
        mtmpItem.内容 = ""
        Call SetChartStyleAndData(Chart, mtmpItem)
        Exit Sub
    End If
    
    '重新显示可用字段
    cboFX.Clear: cboFY.Clear: cboFS.Clear '清除不会激活Click
    strField = mobjDatas("_" & cboData.Text).字段
    If strField <> "" Then
        arrField = Split(strField, "|")
        For i = 0 To UBound(arrField)
            strField = Split(arrField(i), ",")(0)
            Select Case Val(Split(arrField(i), ",")(1))
                Case adDBTimeStamp, adDBTime, adDBDate, adDate
                    cboFX.AddItem strField
                Case adNumeric, adVarNumeric, adSingle, adDouble, adCurrency, adDecimal, adBigInt, adInteger, adSmallInt, adTinyInt, adUnsignedBigInt, adUnsignedInt, adUnsignedSmallInt, adUnsignedTinyInt
                    cboFX.AddItem strField
                    cboFY.AddItem strField
                Case adChar, adVarChar, adLongVarChar, adWChar, adVarWChar, adLongVarWChar, adBSTR
                    cboFS.AddItem strField
            End Select
        Next
    End If
            
    '根据对象值定位字段
    Call GetChartDataName(mtmpItem.内容, strFX, strFS, strFY)
    If strFX <> "" Then
        i = GetCboIndex(cboFX, strFX)
        Call CboSetIndex(cboFX.hWnd, i)
    End If
    If strFS <> "" Then
        i = GetCboIndex(cboFS, strFS)
        Call CboSetIndex(cboFS.hWnd, i)
    End If
    If strFY <> "" Then
        i = GetCboIndex(cboFY, strFY)
        Call CboSetIndex(cboFY.hWnd, i)
    End If
    
    '根据现在值重设对象
    Call SetChartData
End Sub

Private Sub cboFX_Click()
    Call SetChartData
End Sub

Private Sub cboFS_Click()
    Call SetChartData
End Sub

Private Sub cboFY_Click()
    Call SetChartData
End Sub

Private Sub SetChartData()
'功能：根据当前界面的数据设置,重置Chart示例显示
    Dim strFX As String, strFY As String, strFS As String
    Dim str内容 As String

    strFX = cboFX.Text
    strFS = cboFS.Text
    strFY = cboFY.Text
    If strFX <> "" Then
        str内容 = str内容 & "|" & cboData.Text & "." & strFX
    Else
        str内容 = str内容 & "|"
    End If
    If strFS <> "" Then
        str内容 = str内容 & "|" & cboData.Text & "." & strFS
    Else
        str内容 = str内容 & "|"
    End If
    If strFY <> "" Then
        str内容 = str内容 & "|" & cboData.Text & "." & strFY
    Else
        str内容 = str内容 & "|"
    End If
    str内容 = Mid(str内容, 2)
    
    '如果有变化(如更改目标数据源),则重置图表
    If str内容 <> mtmpItem.内容 Then
        mtmpItem.内容 = str内容
        ItemChange = True
    End If
End Sub

Private Sub cboLocate_Click()
    mtmpItem.对齐 = cboLocate.ListIndex
    ItemChange = True
End Sub

Private Sub cboStyle_Click()
    mtmpItem.序号 = cboStyle.ListIndex
        
    Call SetOptionEnabled
    If Visible Then '部份缺省值
        If chkLine.Enabled And chkLine.Value = 0 Then chkLine.Value = 1
        If chkNode.Enabled And chkNode.Value = 0 Then chkNode.Value = 1
    End If
    
    ItemChange = True
End Sub

Private Sub chkGrid_Click()
    If Visible Then
        mtmpItem.网格 = IIF(chkGrid.Value = 1, 1, 0)
        ItemChange = True
    End If
End Sub

Private Sub chkLine_Click()
    If Visible Then
        mtmpItem.下线 = chkLine.Value = 1
        ItemChange = True
    End If
End Sub

Private Sub chkNode_Click()
    If Visible Then
        mtmpItem.自调 = chkNode.Value = 1
        ItemChange = True
    End If
End Sub

Private Sub chkSample_Click()
    If Visible Then
        cboLocate.Enabled = chkSample.Value = 1
        mtmpItem.分栏 = IIF(chkSample.Value = 1, 2, 1)
        ItemChange = True
    End If
End Sub

Private Sub cmdBack_Click()
    On Error Resume Next
    
    cdg.CancelError = True
    cdg.Flags = &H1 Or &H2
    cdg.Color = mtmpItem.背景
    cdg.ShowColor
    If Err.Number = 0 Then
        mtmpItem.背景 = cdg.Color
        cmdBack.BackColor = cdg.Color
        ItemChange = True
    Else
        Err.Clear
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFont_Click()
    On Error Resume Next
    
    cdg.CancelError = True
    cdg.Flags = &H3 Or &H400 Or &H200 Or &H10000
    
    cdg.FontName = mtmpItem.字体
    cdg.FontSize = mtmpItem.字号
    cdg.FontBold = mtmpItem.粗体
    cdg.FontItalic = mtmpItem.斜体

    cdg.ShowFont
    If Err.Number = 0 Then
        On Error GoTo 0
        mtmpItem.字体 = cdg.FontName
        mtmpItem.字号 = cdg.FontSize
        mtmpItem.粗体 = cdg.FontBold
        mtmpItem.斜体 = cdg.FontItalic
        txtFont.Text = cdg.FontName & "," & cdg.FontSize & IIF(cdg.FontBold, ",粗体", "") & IIF(cdg.FontItalic, ",斜体", "")
        Call SelAll(txtFont)
        txtFont.SetFocus
        ItemChange = True
    Else
        Err.Clear
    End If
End Sub

Private Sub cmdFore_Click()
    On Error Resume Next
    
    cdg.CancelError = True
    cdg.Flags = &H1 Or &H2
    cdg.Color = mtmpItem.前景
    cdg.ShowColor
    If Err.Number = 0 Then
        mtmpItem.前景 = cdg.Color
        cmdFore.BackColor = cdg.Color
        ItemChange = True
    Else
        Err.Clear
    End If
End Sub

Private Sub cmdFontTitle_Click()
    Dim arrFont As Variant
    
    On Error Resume Next
    
    cdg.CancelError = True
    cdg.Flags = &H3 Or &H400 Or &H200 Or &H10000
    
    arrFont = Split(Split(mtmpItem.表头, "|")(1), ",")
    cdg.FontName = arrFont(0)
    cdg.FontSize = Val(arrFont(1))
    cdg.FontBold = Val(arrFont(2)) <> 0
    cdg.FontItalic = Val(arrFont(3)) <> 0

    cdg.ShowFont
    If Err.Number = 0 Then
        On Error GoTo 0
        mtmpItem.表头 = Split(mtmpItem.表头, "|")(0) & "|" & cdg.FontName & "," & cdg.FontSize & "," & IIF(cdg.FontBold, 1, 0) & "," & IIF(cdg.FontItalic, 1, 0)
        txtFontTitle.Text = cdg.FontName & "," & cdg.FontSize & IIF(cdg.FontBold, ",粗体", "") & IIF(cdg.FontItalic, ",斜体", "")
        Call SelAll(txtFontTitle)
        txtFontTitle.SetFocus
        ItemChange = True
    Else
        Err.Clear
    End If
End Sub

Private Sub txtFont_GotFocus()
    SelAll txtFont
End Sub

Private Sub txtFont_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 And cmdFont.Enabled Then
        Call cmdFont_Click
    End If
End Sub

Private Sub txtFontTitle_GotFocus()
    SelAll txtFontTitle
End Sub

Private Sub txtFontTitle_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 And cmdFontTitle.Enabled Then
        Call cmdFontTitle_Click
    End If
End Sub

Private Sub txtTitle_Change()
    Dim arrFont As Variant
    
    cmdFontTitle.Enabled = txtTitle.Text <> ""
    txtFontTitle.Enabled = txtTitle.Text <> ""
    
    If Visible Then
        If txtTitle.Text <> "" Then
            If mtmpItem.表头 = "" Then
                mtmpItem.表头 = txtTitle.Text & "|宋体,9,0,0"
            Else
                mtmpItem.表头 = txtTitle.Text & "|" & Split(mtmpItem.表头, "|")(1)
            End If
        Else
            mtmpItem.表头 = ""
        End If
        If mtmpItem.表头 <> "" Then
            arrFont = Split(Split(mtmpItem.表头, "|")(1), ",")
            txtFontTitle.Text = arrFont(0) & "," & Val(arrFont(1)) & IIF(Val(arrFont(2)) <> 0, ",粗体", "") & IIF(Val(arrFont(3)) <> 0, ",斜体", "")
        Else
            txtFontTitle.Text = ""
        End If
        ItemChange = True
    End If
End Sub

Private Sub txtTitle_GotFocus()
    Call SelAll(txtTitle)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call PressKey(vbKeyTab)
    Else
        If InStr("'|,;", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    Dim strData As String, i As Long
    Dim arrFont As Variant
    
    mblnOK = False
    Call CboSetWidth(cboStyle.hWnd, 3400)
    Call CboSetHeight(cboStyle, Screen.Height)
    Call CopyItem(mtmpItem, mobjItem)
    
    '数据来源
    For i = 1 To mobjDatas.Count
        cboData.AddItem mobjDatas(i).名称
    Next
    If mtmpItem.内容 <> "" Then
        Call GetChartDataName(mtmpItem.内容, , , , strData)
        cboData.ListIndex = GetCboIndex(cboData, strData)
    End If
        
    '图表样式
    cboStyle.AddItem "散点图(单一X,Y坐标序列)"
    cboStyle.AddItem "折线图"
    cboStyle.AddItem "条形图"
    cboStyle.AddItem "饼图"
    cboStyle.AddItem "层叠图"
    cboStyle.AddItem "面积图"
    cboStyle.AddItem "股价图(盘高,盘低)"
    cboStyle.AddItem "股价图(盘高,盘低,开盘,收盘)"
    cboStyle.AddItem "股价图(阴阳烛图:盘高,盘低,开盘,收盘)"
    cboStyle.AddItem "级线图"
    cboStyle.AddItem "雷达图"
    cboStyle.AddItem "填充雷达图"
    cboStyle.AddItem "气泡图"
    Call CboSetIndex(cboStyle.hWnd, mtmpItem.序号)
    
    '标题
    If mtmpItem.表头 <> "" Then
        txtTitle.Text = Split(mtmpItem.表头, "|")(0)
        arrFont = Split(Split(mtmpItem.表头, "|")(1), ",")
        txtFontTitle.Text = arrFont(0) & "," & Val(arrFont(1)) & IIF(Val(arrFont(2)) <> 0, ",粗体", "") & IIF(Val(arrFont(3)) <> 0, ",斜体", "")
    End If
            
    '图表字体
    txtFont.Text = mtmpItem.字体 & "," & mtmpItem.字号 & IIF(mtmpItem.粗体, ",粗体", "") & IIF(mtmpItem.斜体, ",斜体", "")
            
    '图表颜色
    cmdFore.BackColor = mtmpItem.前景
    cmdBack.BackColor = mtmpItem.背景
    
    '图例
    chkSample.Value = IIF(mtmpItem.分栏 <= 1, 0, 1)
    cboLocate.Enabled = chkSample.Value = 1
    cboLocate.AddItem "1-右面"
    cboLocate.AddItem "2-下面"
    cboLocate.AddItem "3-左面"
    cboLocate.AddItem "4-上面"
    cboLocate.AddItem "5-右下角"
    cboLocate.AddItem "6-左下角"
    'cboLocate.AddItem "7-右上角"
    'cboLocate.AddItem "8-左上角"
    Call CboSetIndex(cboLocate.hWnd, mtmpItem.对齐)
        
    '其它格式：数字位串,三维效果|XY轴互换
    '三维效果
    chkFormat(0).Value = IIF(Val(Mid(Format(mtmpItem.格式, "00"), 1, 1)) = 0, 0, 1)
    'XY轴互换
    chkFormat(1).Value = IIF(Val(Mid(Format(mtmpItem.格式, "00"), 2, 1)) = 0, 0, 1)
        
    '其它：
    chkGrid.Value = IIF(mtmpItem.网格 <> 0, 1, 0)
    chkLine.Value = IIF(mtmpItem.下线, 1, 0)
    chkNode.Value = IIF(mtmpItem.自调, 1, 0)
                
    '设置可选项
    Call SetOptionEnabled
    
    ItemChange = False
    Call SetChartStyleAndData(Chart, mtmpItem)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mtmpItem = Nothing
End Sub

Private Function CheckInput() As Boolean
    If cboFX.Text = "" Then
        MsgBox "请指定Ｘ值字段来源。", vbInformation, App.Title
        cboFX.SetFocus: Exit Function
    End If
    If cboFS.Text = "" Then
        MsgBox "请指定序列字段来源。", vbInformation, App.Title
        cboFS.SetFocus: Exit Function
    End If
    If cboFY.Text = "" Then
        MsgBox "请指定Ｙ值字段来源。", vbInformation, App.Title
        cboFY.SetFocus: Exit Function
    End If
    If cboFX.Text = cboFY.Text Then
        MsgBox "Ｙ值字段与Ｘ值字段不能相同。", vbInformation, App.Title
        cboFY.SetFocus: Exit Function
    End If
    CheckInput = True
End Function
