VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOpenStudyList 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "打开检查"
   ClientHeight    =   5910
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   12090
   Icon            =   "frmOpenStudyList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   12090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picPanel 
      Align           =   2  'Align Bottom
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   12090
      TabIndex        =   0
      Top             =   4935
      Width           =   12090
      Begin VB.CommandButton cmdCancel 
         Caption         =   "取 消(&S)"
         Height          =   375
         Left            =   10800
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdSure 
         Caption         =   "确 定(&S)"
         Height          =   375
         Left            =   9120
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
   End
   Begin zl9PACSWork.ucFlexGrid ufgStudyList 
      Height          =   3975
      Left            =   600
      TabIndex        =   3
      Top             =   600
      Width           =   10935
      _ExtentX        =   21405
      _ExtentY        =   8705
      HeadCheckValue  =   1
      IsCopyAdoMode   =   0   'False
      IsEjectConfig   =   -1  'True
      Editable        =   0
      ReadOnly        =   -1  'True
      IsShowPopupMenu =   0   'False
      HeadFontCharset =   134
      HeadFontWeight  =   400
      HeadColor       =   0
      DataFontCharset =   134
      DataFontWeight  =   400
      DataColor       =   -2147483640
      GridLineColor   =   14737632
   End
   Begin MSComctlLib.ImageList Imglist 
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
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOpenStudyList.frx":000C
            Key             =   "紧急"
            Object.Tag             =   "1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOpenStudyList.frx":05A6
            Key             =   "住院"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOpenStudyList.frx":0E80
            Key             =   "阳性"
            Object.Tag             =   "3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOpenStudyList.frx":0FDA
            Key             =   "影像"
            Object.Tag             =   "4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOpenStudyList.frx":1754
            Key             =   "绿色通道"
            Object.Tag             =   "6"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOpenStudyList.frx":18AE
            Key             =   "路径"
            Object.Tag             =   "7"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOpenStudyList.frx":1E48
            Key             =   "无费"
            Object.Tag             =   "7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmOpenStudyList.frx":21E2
            Key             =   "收费"
            Object.Tag             =   "8"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmOpenStudyList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mrsData As ADODB.Recordset

Public mlngModule As Long
Public blnOk As Boolean
Public mblncmd已缴 As Boolean, mblncmd未缴 As Boolean, mblncmd无费 As Boolean, mblncmd补缴 As Boolean

Private mlngTempCharged As Long


Public Sub ShowStudyWindow(ByVal Cols As String, rsData As ADODB.Recordset, owner As Object, imgList As ImageList)
'显示检查窗口
    Dim strFilter As String
    
    Set mrsData = rsData
        
    '只显示检查过程为报到2，检查3，报告中4的检查数据
    strFilter = "检查过程=2 or 检查过程=3 or 检查过程=4"

    Set ufgStudyList.ImageList = imgList
    
    ufgStudyList.ColNames = Replace(Cols, "btn,", "")   '在该列表中，不需要按钮
    ufgStudyList.ColConvertFormat = ""
    ufgStudyList.DefaultColNames = ""
    ufgStudyList.IsKeepRows = False
        
    Set ufgStudyList.AdoData = mrsData
    ufgStudyList.AdoFilter = strFilter
    
    Call ufgStudyList.BindData
    
    Call ufgStudyList.HidenCol(ufgStudyList.GetColIndex("路径"))
    Call ufgStudyList.HidenCol(ufgStudyList.GetColIndex("阳性"))
    Call ufgStudyList.HidenCol(ufgStudyList.GetColIndex("危急"))
    Call ufgStudyList.HidenCol(ufgStudyList.GetColIndex("报告质量"))
    Call ufgStudyList.HidenCol(ufgStudyList.GetColIndex("报告打印"))
    Call ufgStudyList.HidenCol(ufgStudyList.GetColIndex("报告发放"))
    
    
    
    If mlngModule = G_LNG_PATHOLSYS_NUM Then    '获取病理检查执行状态
        Call ufgStudyList.HidenCol(ufgStudyList.GetColIndex("病理执行状态"))
        Call ufgStudyList.HidenCol(ufgStudyList.GetColIndex("质量"))
    Else
        Call ufgStudyList.HidenCol(ufgStudyList.GetColIndex("胶片打印"))
        Call ufgStudyList.HidenCol(ufgStudyList.GetColIndex("符合情况"))
        Call ufgStudyList.HidenCol(ufgStudyList.GetColIndex("影像质量"))
    End If
    
    Call ufgStudyList.LocateRow(1)
    
    '显示检查列表窗口
    Call Me.Show(1, owner)
End Sub

Private Sub cmdCancel_Click()
    blnOk = False
    Call Me.Hide
End Sub

Private Sub cmdSure_Click()
    If Not ufgStudyList.IsSelectionRow Then
        Call MsgBoxD(Me, "请选择需要进行图像采集的检查记录。", vbOKOnly, gstrSysName)
        Exit Sub
    End If
    
    blnOk = True
    Call Me.Hide
End Sub

Private Sub Form_Load()
    SetWindowPos Me.hWnd, -1, Me.CurrentX, Me.CurrentY, Me.ScaleWidth, Me.ScaleHeight, 3 '将窗口置顶
    
    Call RestoreWinState(Me, App.ProductName)
    
    blnOk = False
End Sub

Private Sub Form_Resize()
On Error GoTo ErrHandle
    ufgStudyList.Left = 120
    ufgStudyList.Top = 120
    ufgStudyList.Height = Me.ScaleHeight - picPanel.Height - 240
    ufgStudyList.Width = Me.ScaleWidth - 240
    
    cmdCancel.Left = picPanel.Width - cmdCancel.Width - 120
    cmdSure.Left = cmdCancel.Left - cmdSure.Width - 120
    
    Exit Sub
ErrHandle:
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub ufgStudyList_DblClick()
    If Val(ufgStudyList.CurKeyValue) > 0 Then
        blnOk = True
        Me.Hide
    End If
End Sub



Private Sub ufgStudyList_OnFilterRowData(rsData As ADODB.Recordset, rsClone As ADODB.Recordset, blnFilterOut As Boolean)
    '判断是否已经收费
    '"病人医嘱发送.记录性质"--- 1是收费的，2是记帐的。
    
    '通过"病人医嘱发送.计费状态"直接判断,原有值：-1-无须计费;0-未计费;1-已计费，对于记帐单（包括门诊记帐单），保持原有值不变。
    '对于收费单的发送记录，增加两种状态：2-部分收费，3-全部收费
    
    '没有对应费用的医嘱有两种情况，一种是"-1-无须计费"，即没有设置收费对照，一种是"0-未计费"，即虽然设置了收费对照，但设置为发送后手工计费，即在医技科室去生成。
    '"1-已计费"就是发送时生成了费用的。但生成了费用单据不表示收费了，生成可能是记帐划价单，或收费划价单，其中收费划价单就多两种状态。
    '"2-部分收费"表示部分收费和部分退费的情况，反正没收得完。
    
    '已收费显示状态：已收费；无费用；未收费：
    '未收费----
    '1、主医嘱是收费单的，满足以下条件算未收费
    '   (1)有一条主医嘱和部位医嘱的 计费状态 in (1,2)算未收费 ------“记录性质=1 and 计费状态 in (1,2)”
    '已收费：
    '1、主医嘱是记账的算收费-------“记录性质=2”
    '2、主医嘱是收费单的，满足以下条件算收费
    '   (1)排除未收费后，有一条主医嘱和部位医嘱的 计费状态 =3 算收费-----“记录性质=1 and 计费状态 = 3”
    '无费用
    '1、主医嘱是收费单的，满足以下条件算无费用
    '   (1)所有主医嘱和部位医嘱的 计费状态 in (-1,0)算无费用 ------“记录性质=1 and 计费状态 in (-1,0)”
    
    
    ' intCharged  '0--未收费；1--已收费；2--无费用
    
    If Nvl(rsData!相关ID) <> "" Then
        '相关id不为空时，说明书部位医嘱，不需要显示到列表中
        blnFilterOut = True
        Exit Sub
    End If

    mlngTempCharged = 2 '无费用
    
    If Nvl(rsData!记录性质, 2) = 2 Then
        '住院登记的病人，如果没有计费，则归为无费用
        If Nvl(rsData!计费状态, -1) = 0 Then
            mlngTempCharged = 2
        Else
            mlngTempCharged = 1  '已收费
        End If
    Else
        If Nvl(rsData!计费状态, -1) = 1 Or Nvl(rsData!计费状态, -1) = 2 Then
            mlngTempCharged = 0      '未收费
        Else        '主医嘱的计费状态是 -1,0,3  （3--已收费；-1，0--无费用）
            '查询主医嘱未计费或者已经收费了，还要查部位医嘱的收费情况，所有医嘱都已经收费，才算是收费
            
            '如果主费用是已收费的，先记录成已收费
            If Nvl(rsData!计费状态, -1) = 3 Then
                mlngTempCharged = 1      '已收费
            End If
            
            rsClone.Filter = "相关ID = " & Nvl(rsData!医嘱ID)
            Do While rsClone.EOF = False
                If Nvl(rsClone!计费状态, -1) = 1 Or Nvl(rsClone!计费状态, -1) = 2 Then
                    mlngTempCharged = 0      '未收费

                    Exit Do
                ElseIf Nvl(rsClone!计费状态, -1) = 3 Then
                    mlngTempCharged = 1      '已收费
                End If

                rsClone.MoveNext
            Loop
            
'            '计费状态：-1-无须计费(通常无执行和院外执行的都无须计费);0-未计费;1-已计费，对收费单据多两种状态:2-部分收费，3-全部收费
'            rsClone.Filter = "相关ID = " & Nvl(rsData!医嘱ID) & " and 计费状态=1 and 计费状态=2"
'            If rsClone.RecordCount > 0 Then
'                mlngTempCharged = 0 '未收费
'            Else
'                rsClone.Filter = "相关ID = " & Nvl(rsData!医嘱ID) & " and 计费状态=3"
'                If rsClone.RecordCount > 0 Then mlngTempCharged = 1 '已收费
'            End If
            
        End If
    End If

    If mlngModule = G_LNG_PATHOLSYS_NUM Then
        If Nvl(rsData!补费) > 0 Then mlngTempCharged = 4 '需要补费，需补费的检查也是未收费的检查
    End If
    
    If Nvl(rsData!相关ID) = "" And ((mblncmd已缴 = True And mlngTempCharged = 1) Or (mblncmd未缴 = True And (mlngTempCharged = 0 Or mlngTempCharged = 4)) _
        Or (mblncmd无费 = True And mlngTempCharged = 2) Or (mblncmd补缴 = True And mlngTempCharged = 4) _
        Or (mblncmd已缴 = False And mblncmd未缴 = False And mblncmd补缴 = False And mblncmd无费 = False)) Then
        blnFilterOut = False
        
        Call RowDataConvert(rsData)
    Else
        blnFilterOut = True
    End If
End Sub


Private Sub RowDataConvert(rsData As ADODB.Recordset)
    Dim rsBaby As ADODB.Recordset
    Dim intTxtLen As Long
    
    '如果该数据要显示，则需要转换数据中的部分值
    rsData!申请单 = IIf(Nvl(rsData!申请单) = "", "无", "已扫描")
    rsData!检查过程 = IIf(Val(Nvl(rsData!执行状态)) = 2, "已拒绝", Decode(Val(Nvl(rsData!检查状态, 0)), -1, "已驳回", 0, "已登记", 1, "已登记", _
                                                                                2, IIf(Nvl(rsData!报告操作) <> "", "处理中", _
                                                                                        IIf(Nvl(rsData!报告人) = "", "已报到", "报告中")), _
                                                                                3, IIf(Nvl(rsData!报告操作) <> "", "处理中", _
                                                                                        IIf(Nvl(rsData!报告人) = "", "已检查", "报告中")), _
                                                                                4, IIf(Nvl(rsData!报告操作) <> "", "处理中", _
                                                                                        IIf(Nvl(rsData!复核人) <> "", "审核中", "已报告")), _
                                                                                5, "已审核", "已完成"))
                                                                                
    If Nvl(rsData!婴儿) <> 0 Then
        gstrSQL = "Select Nvl(A.婴儿姓名, B.姓名 || '之子' || Trim(To_Char(A.序号, '9'))) As 婴儿姓名, 婴儿性别, 出生时间" & vbNewLine & _
                    "From 病人新生儿记录 A, 病人信息 B" & vbNewLine & _
                    "Where A.病人id = [1] And A.主页id = [2] And A.病人id = B.病人id And A.序号 = [3]"
        
        Set rsBaby = zlDatabase.OpenSQLRecord(gstrSQL, "提取婴儿信息", CLng(rsData!病人ID), CLng(Nvl(rsData!主页ID, 0)), CLng(rsData!婴儿))
        
        If Not rsBaby.EOF Then
            rsData!姓名 = rsBaby!婴儿姓名
            rsData!性别 = Nvl(rsBaby!婴儿性别)
            rsData!年龄 = Nvl(rsBaby!出生时间)
        End If
    End If
    
    
    If InStr(Nvl(rsData!医嘱内容), ":") > 0 Then '新的模式保存在医嘱内容中信息是 名称,执行标记:部位(方法,方法),部位---
        rsData!部位方法 = Split(Nvl(rsData!医嘱内容), ":")(1)
        rsData!医嘱内容 = Split(Nvl(rsData!医嘱内容), ":")(0)
    End If
    
    
    If Val(Nvl(rsData!紧急)) <> 0 Then
        rsData!紧急 = " "
    Else
        rsData!紧急 = ""
    End If
    
    If mlngTempCharged = 0 Then  '未收费
        rsData!收费 = ""
    ElseIf mlngTempCharged = 1 Then   '已收费
        rsData!收费 = " "
    ElseIf mlngTempCharged = 2 Then    '无费用
        rsData!收费 = "  "
    Else
        rsData!收费 = "   "
    End If
    
    If rsData!来源 = 1 Then
        rsData!来源 = "门"
    ElseIf rsData!来源 = 2 Then
        rsData!来源 = "住"
    ElseIf rsData!来源 = 3 Then
        rsData!来源 = "外"
    ElseIf rsData!来源 = 4 Then
        rsData!来源 = "体"
    End If
End Sub


Private Sub ufgStudyList_OnRefreshRowData(rsBind As ADODB.Recordset, ByVal lngRow As Long)
On Error GoTo ErrHandle
    Dim strTag As String
    Dim strTemp As String
    Dim i As Long
    
    For i = 0 To ufgStudyList.DataGrid.Cols - 1
        Select Case ufgStudyList.DataGrid.TextMatrix(0, i)
                
                
            Case "紧急"
                If ufgStudyList.Text(lngRow, "紧急") = " " Then
                    Set ufgStudyList.DataGrid.Cell(flexcpPicture, lngRow, i) = imgList.ListImages("紧急").Picture
                End If
        
            Case "来源"
                strTag = Decode(ufgStudyList.Text(lngRow, "来源"), "门", 1, "住", 2, "外", 3, 4)
                ufgStudyList.DataGrid.Cell(flexcpData, lngRow, i) = strTag
                
                If ufgStudyList.Text(lngRow, "来源") = "住" Then
                    Set ufgStudyList.DataGrid.Cell(flexcpPicture, lngRow, i) = imgList.ListImages("住院").Picture
                End If
                
            Case "收费" 'TODO:病理还需要考虑补缴费用的情况
                If ufgStudyList.Text(lngRow, "收费") = "" Then  '未收费
                    Set ufgStudyList.DataGrid.Cell(flexcpPicture, lngRow, i) = imgList.ListImages("无费").Picture
                ElseIf ufgStudyList.Text(lngRow, "收费") = " " Then   '已收费
                    Set ufgStudyList.DataGrid.Cell(flexcpPicture, lngRow, i) = imgList.ListImages("收费").Picture
                ElseIf ufgStudyList.Text(lngRow, "收费") = "   " Then   '补费
                    Set ufgStudyList.DataGrid.Cell(flexcpPicture, lngRow, i) = imgList.ListImages("补费").Picture
                Else '无费用("  ")
                    '无费用不显示图标
                End If

                
            Case "姓名" '如果为绿色通道，则需要在姓名面前添加图标
                If Val(ufgStudyList.Text(lngRow, "绿色通道")) <> 0 Then
                    Set ufgStudyList.DataGrid.Cell(flexcpPicture, lngRow, i) = imgList.ListImages("绿色通道").Picture
                End If
                
            Case GetStudyNumberDisplayName  '检查号或者病理号
                If ufgStudyList.Text(lngRow, "检查UID") <> "" Then
                    '病理系统中，检查列表中的检查号显示为病理号
                    Set ufgStudyList.DataGrid.Cell(flexcpPicture, lngRow, i) = imgList.ListImages(IIf(mlngModule = G_LNG_PATHOLSYS_NUM, "病理", "影像")).Picture
                End If
                
            Case "检查过程"
                '根据检查过程，设置不同的颜色
                If ufgStudyList.Text(lngRow, "检查过程") = "已拒绝" Then ufgStudyList.DataGrid.Cell(flexcpBackColor, lngRow, 1, lngRow, ufgStudyList.DataGrid.Cols - 1) = gdblColor已拒绝
                If ufgStudyList.Text(lngRow, "检查过程") = "已完成" Then ufgStudyList.DataGrid.Cell(flexcpBackColor, lngRow, 1, lngRow, ufgStudyList.DataGrid.Cols - 1) = gdblColor已完成
                If ufgStudyList.Text(lngRow, "检查过程") = "已报到" Then ufgStudyList.DataGrid.Cell(flexcpBackColor, lngRow, 1, lngRow, ufgStudyList.DataGrid.Cols - 1) = gdblColor已报到
                If ufgStudyList.Text(lngRow, "检查过程") = "已登记" Then ufgStudyList.DataGrid.Cell(flexcpBackColor, lngRow, 1, lngRow, ufgStudyList.DataGrid.Cols - 1) = gdblColor已登记
                If ufgStudyList.Text(lngRow, "检查过程") = "已检查" Then ufgStudyList.DataGrid.Cell(flexcpBackColor, lngRow, 1, lngRow, ufgStudyList.DataGrid.Cols - 1) = gdblColor已检查
                If ufgStudyList.Text(lngRow, "检查过程") = "已审核" Then ufgStudyList.DataGrid.Cell(flexcpBackColor, lngRow, 1, lngRow, ufgStudyList.DataGrid.Cols - 1) = gdblColor已审核
                If ufgStudyList.Text(lngRow, "检查过程") = "处理中" Then ufgStudyList.DataGrid.Cell(flexcpBackColor, lngRow, 1, lngRow, ufgStudyList.DataGrid.Cols - 1) = gdblColor处理中
                If ufgStudyList.Text(lngRow, "检查过程") = "报告中" Then ufgStudyList.DataGrid.Cell(flexcpBackColor, lngRow, 1, lngRow, ufgStudyList.DataGrid.Cols - 1) = gdblColor报告中
                If ufgStudyList.Text(lngRow, "检查过程") = "审核中" Then ufgStudyList.DataGrid.Cell(flexcpBackColor, lngRow, 1, lngRow, ufgStudyList.DataGrid.Cols - 1) = gdblColor审核中
                If ufgStudyList.Text(lngRow, "检查过程") = "已报告" Then ufgStudyList.DataGrid.Cell(flexcpBackColor, lngRow, 1, lngRow, ufgStudyList.DataGrid.Cols - 1) = gdblColor已报告
                If ufgStudyList.Text(lngRow, "检查过程") = "已驳回" Then ufgStudyList.DataGrid.Cell(flexcpBackColor, lngRow, 1, lngRow, ufgStudyList.DataGrid.Cols - 1) = gdblColor已驳回
                                
        End Select
        
    Next i
    
ErrHandle:
    Exit Sub
End Sub


Private Function GetStudyNumberDisplayName() As String
'获取检查号码显示名称
    GetStudyNumberDisplayName = IIf(mlngModule = G_LNG_PATHOLSYS_NUM, "病理号", "检查号")
End Function
