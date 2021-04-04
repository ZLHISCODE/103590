VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm保险项目选择奉庆 
   AutoRedraw      =   -1  'True
   Caption         =   "医保项目选择"
   ClientHeight    =   5490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7845
   Icon            =   "frm保险项目选择奉庆.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   7845
   Begin VB.PictureBox picBack 
      BackColor       =   &H8000000A&
      Height          =   600
      Left            =   0
      ScaleHeight     =   540
      ScaleWidth      =   7785
      TabIndex        =   10
      Top             =   4890
      Visible         =   0   'False
      Width           =   7845
      Begin MSComctlLib.ProgressBar prgs 
         Height          =   450
         Left            =   795
         TabIndex        =   11
         Top             =   45
         Width           =   6240
         _ExtentX        =   11007
         _ExtentY        =   794
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lblInfor 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "费用类别"
         Height          =   255
         Left            =   60
         TabIndex        =   12
         Top             =   180
         Width           =   720
      End
   End
   Begin VB.PictureBox picSplit 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   930
      Left            =   2340
      MousePointer    =   9  'Size W E
      ScaleHeight     =   930
      ScaleWidth      =   45
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1575
      Width           =   45
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshGrid 
      Height          =   3990
      Left            =   3045
      TabIndex        =   6
      Top             =   390
      Width           =   4680
      _ExtentX        =   8255
      _ExtentY        =   7038
      _Version        =   393216
      FixedCols       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      BandDisplay     =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   45
      Top             =   3900
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
            Picture         =   "frm保险项目选择奉庆.frx":0E42
            Key             =   "Detail"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm保险项目选择奉庆.frx":1C94
            Key             =   "Class"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwClass 
      Height          =   4050
      Left            =   0
      TabIndex        =   7
      Top             =   255
      Width           =   2970
      _ExtentX        =   5239
      _ExtentY        =   7144
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "img16"
      Appearance      =   1
   End
   Begin VB.PictureBox picCmd 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   510
      Left            =   0
      ScaleHeight     =   510
      ScaleWidth      =   7845
      TabIndex        =   1
      Top             =   4980
      Width           =   7845
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         Default         =   -1  'True
         Height          =   350
         Left            =   5400
         TabIndex        =   5
         Top             =   180
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   6660
         TabIndex        =   4
         Top             =   180
         Width           =   1100
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "打印列表"
         Height          =   350
         Left            =   15
         TabIndex        =   3
         Top             =   180
         Width           =   1100
      End
      Begin VB.CommandButton cmdRequery 
         Caption         =   "项目下载"
         Height          =   350
         Left            =   1335
         TabIndex        =   2
         ToolTipText     =   "从中心下载服务项目、病种信息和定点医疗机构"
         Top             =   180
         Width           =   1100
      End
   End
   Begin VB.Label lblClass 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "项目大类(&K)"
      Height          =   240
      Left            =   15
      TabIndex        =   9
      Top             =   0
      Width           =   2970
   End
   Begin VB.Label lblDetail 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "项目明细(&D)"
      Height          =   240
      Left            =   3060
      TabIndex        =   8
      Top             =   15
      Width           =   4710
   End
End
Attribute VB_Name = "frm保险项目选择奉庆"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mint险类 As Integer
Private mstrCode As String
Private mstrName As String
Private mblnOK As Boolean

Private mLocalCode As String '指向编码
Private mblnFirst As Boolean
Private mbln诊疗 As Boolean

Private Sub cmdCancel_Click()
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Trim(mshGrid.TextMatrix(mshGrid.Row, 0)) = "" Then
        MsgBox "没有选择项目！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '返回选择项目编码
    mstrCode = mshGrid.TextMatrix(mshGrid.Row, 0) & Trim(mshGrid.TextMatrix(mshGrid.Row, 1))
    mstrName = mshGrid.TextMatrix(mshGrid.Row, 2)
    mblnOK = True
    Unload Me
End Sub

Private Function Loadtree() As Boolean
    Dim rsTemp As New ADODB.Recordset, strTemp As String
    Dim tmpNode As Node
    mblnOK = False
    
    On Error GoTo ErrHand:
    
    '装载数据
    '
    tvwClass.Nodes.Clear
    Set tmpNode = tvwClass.Nodes.Add(, 4, "K1", "【1】药品", "Detail", "Detail")
    tmpNode.Sorted = True
    tmpNode.Selected = True
    
    Set tmpNode = tvwClass.Nodes.Add(, 4, "K2", "【2】诊疗", "Detail", "Detail")
    tmpNode.Sorted = True
    
    Set tmpNode = tvwClass.Nodes.Add(, 4, "K3", "【4】服务", "Detail", "Detail")
    tmpNode.Sorted = True
    
    'Call FillList
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2
    Call RestoreWinState(Me, App.ProductName)
    Loadtree = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Loadtree = False
End Function
Public Function GetCode(ByVal frmMain As Form, strCode As String, strName As String, Optional bln诊疗 As Boolean = False) As Boolean
    '功能：获取编码
    '参数：strCode-编码(类别+编码)
    '返回：成功返回True
    mLocalCode = strCode
    frm保险项目选择奉庆.Show vbModal, frm保险项目
    
    '返回值
    If mblnOK = True Then
        strCode = mstrCode
        strName = mstrName
    End If
    GetCode = mblnOK
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub SetGrdColHead(Optional ByVal blnInit As Boolean = True)
    With mshGrid
        .Redraw = False
        If blnInit Then
            .Clear
            .Rows = 2
            .Cols = 15
            .TextMatrix(0, 0) = "类别"
            .TextMatrix(0, 1) = "编码"
            .TextMatrix(0, 2) = "名称"
            .TextMatrix(0, 3) = "英文名称"
            .TextMatrix(0, 4) = "收费类别"
            .TextMatrix(0, 5) = "收费等级"
            .TextMatrix(0, 6) = "助记码"
            .TextMatrix(0, 7) = "单位"
            .TextMatrix(0, 8) = "标准价格"
            .TextMatrix(0, 9) = "支付标准"
            .TextMatrix(0, 10) = "剂型"
            .TextMatrix(0, 11) = "规格"
            .TextMatrix(0, 12) = "备注"
            .TextMatrix(0, 13) = "变更时间"
            .TextMatrix(0, 14) = "维护标志"
        End If
        .ColWidth(0) = 0
        .ColWidth(1) = 1500
        .ColWidth(2) = 2000
        .ColWidth(3) = 1500
        .ColWidth(4) = 1000
        .ColWidth(5) = 1000
        .ColWidth(6) = 1000
        .ColWidth(8) = 1000
        .ColWidth(9) = 1000
        .ColWidth(10) = 1400
        .ColWidth(11) = 1400
        .ColWidth(12) = 2000
        .ColWidth(13) = 1600
        .ColWidth(14) = 1000
        
        .ColAlignment(0) = 0
        .ColAlignment(1) = 1
        .ColAlignment(2) = 1
        .ColAlignment(3) = 1
        .ColAlignment(4) = 4
        .ColAlignment(5) = 4
        .ColAlignment(6) = 4
        .ColAlignment(8) = 4
        .ColAlignment(9) = 7
        .ColAlignment(10) = 7
        .ColAlignment(11) = 1
        .ColAlignment(12) = 1
        .ColAlignment(13) = 4
        .ColAlignment(14) = 4
        .Redraw = True
End With

End Sub
Private Sub FillList()
    '功能：显示当前类别下的医保明细
    Dim rsTemp As New ADODB.Recordset
    Dim lst As ListItem, fld As ADODB.Field
    Dim str类别代码 As String, blnColSet As Boolean
    Dim lngCol  As Long
    Dim varValue As Variant
    
    Me.MousePointer = vbHourglass
    
    On Error GoTo errHandle
    With tvwClass.SelectedItem
        str类别代码 = Mid(.Key, 2)
    End With
    
    
    rsTemp.CursorLocation = adUseClient
    gstrSQL = " select  类别,编码,名称,英文名称,收费类别,收费等级,助记码,单位,标准价格,支付标准,剂型,规格,备注,变更时间,维护标志 " & _
             "  from 医保收费目录" & _
             "  where 类别=" & Val(str类别代码)
    
    rsTemp.Open gstrSQL, gcnOracle_奉庆, adOpenStatic, adLockReadOnly
    If rsTemp.RecordCount = 0 Then
        '设置列头
        Call SetGrdColHead
    Else
        Set mshGrid.DataSource = rsTemp
        Call SetGrdColHead(False)
    End If
    Me.MousePointer = vbDefault
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Me.MousePointer = vbDefault
End Sub
Private Sub cmdPrint_Click()
    If gstrUserName = "" Then Call GetUserInfo
    subPrint 1
End Sub

Private Sub subPrint(bytMode As Byte)
    '功能:进行打印,预览和输出到EXCEL
    '参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    Dim objPrint As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    Dim nod As Node
    
    Set nod = tvwClass.SelectedItem
    Set objPrint.Body = mshGrid
    objPrint.Title.Text = "保险项目"
    
    objRow.Add "医保大类：" & nod.Text
    objPrint.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "打印人：" & gstrUserName
    objRow.Add "打印时间：" & Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    objPrint.BelowAppRows.Add objRow
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrView1Grd objPrint, 1
          Case 2
              zlPrintOrView1Grd objPrint, 2
          Case 3
              zlPrintOrView1Grd objPrint, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub
Private Sub cmdRequery_Click()
    Dim strInput As String
    Dim rsTemp As New ADODB.Recordset
    Dim bln病种 As Boolean
    
    If MsgBox("本操作可能会花比较长的时间，是否继续？" & vbCrLf & vbCrLf & "另外注意，本操作只更新医保项目明细，而不包括对应关系。", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Exit Sub
    End If
        
    MousePointer = vbHourglass
 
    picCmd.Enabled = False
    tvwClass.Enabled = False
        
    cmdRequery.Visible = False
    cmdCancel.Enabled = False
    cmdPrint.Visible = False
        
    With picBack
        .Left = 0
        .Width = ScaleWidth
        .Top = ScaleHeight - .Height
        picBack.Visible = True
    End With
    
    '下载服务项目目录
    '1.下载药品
    lblInfor.Caption = "药品"
    If 下载服务项目目录_奉庆(1, prgs) = False Then
        GoTo GoEnd:
    End If
    '2.下载诊疗
    lblInfor.Caption = "诊疗"
    If 下载服务项目目录_奉庆(2, prgs) = False Then
        GoTo GoEnd:
    End If
    '3.下载服务
    lblInfor.Caption = "服务"
    If 下载服务项目目录_奉庆(3, prgs) = False Then
        GoTo GoEnd:
    End If
    '4.下载费用类别
    lblInfor.Caption = "费用类别"
    If 下载服务项目目录_奉庆(4, prgs) = False Then
       GoTo GoEnd:
    End If
    '5.下载病种
    lblInfor.Caption = "病种"
    If 下载服务项目目录_奉庆(5, prgs) = False Then
        GoTo GoEnd:
    End If
GoEnd:
    MousePointer = vbDefault
    picCmd.Enabled = True
    tvwClass.Enabled = True
    picBack.Visible = False
    cmdRequery.Visible = True
    cmdCancel.Enabled = True
    cmdPrint.Visible = True
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    If Loadtree = False Then
        Exit Sub
    End If
End Sub
Private Sub Form_Load()
    mblnFirst = True
End Sub

Private Sub Form_Resize()
    lblClass.Top = 0: lblClass.Left = 0: lblClass.Width = tvwClass.Width
    
    On Error Resume Next
    
    tvwClass.Left = 0: tvwClass.Top = lblClass.Top + lblClass.Height
    tvwClass.Height = Me.ScaleHeight - lblClass.Height - picCmd.Height
    
    picSplit.Top = tvwClass.Top
    picSplit.Left = tvwClass.Left + tvwClass.Width
    picSplit.Height = tvwClass.Height
    
    lblDetail.Top = lblClass.Top
    If tvwClass.Visible = True Then
        lblDetail.Left = picSplit.Left + picSplit.Width
    Else
        lblDetail.Left = 0
    End If
    lblDetail.Width = Me.ScaleWidth - lblDetail.Left
    With picBack
        .Left = 0
        .Width = ScaleWidth
    End With
    With mshGrid
        .Top = tvwClass.Top
        .Left = lblDetail.Left
        .Width = lblDetail.Width
        .Height = tvwClass.Height
    End With
End Sub

Private Sub picBack_Resize()
    Err = 0
    On Error Resume Next
    With prgs
        .Left = lblInfor.Left + lblInfor.Width
        .Width = picBack.ScaleWidth - .Left
    End With
End Sub

Private Sub picCmd_Resize()
    cmdCancel.Left = picCmd.ScaleWidth - cmdCancel.Width * 1.4
    cmdOK.Left = cmdCancel.Left - cmdOK.Width * 1.25
    cmdPrint.Top = cmdOK.Top
    cmdRequery.Top = cmdOK.Top
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
End Sub

Private Sub mshgrid_DblClick()
    cmdOK_Click
End Sub

Private Sub picSplit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If tvwClass.Width + x < 1000 Or mshGrid.Width - x < 1000 Then Exit Sub
        picSplit.Left = picSplit.Left + x
        lblClass.Width = lblClass.Width + x
        tvwClass.Width = tvwClass.Width + x
        
        lblDetail.Left = lblDetail.Left + x
        lblDetail.Width = lblDetail.Width - x
        
        mshGrid.Left = mshGrid.Left + x
        mshGrid.Width = mshGrid.Width - x
    End If
End Sub

Private Sub tvwClass_NodeClick(ByVal Node As MSComctlLib.Node)
    Call FillList
End Sub







