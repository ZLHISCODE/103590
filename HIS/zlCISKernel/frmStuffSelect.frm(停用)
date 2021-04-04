VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStuffSelect 
   AutoRedraw      =   -1  'True
   Caption         =   "备货卫材选择器"
   ClientHeight    =   6705
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10185
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmStuffSelect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   10185
   Begin VB.PictureBox picStuffItem 
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   3255
      ScaleHeight     =   480
      ScaleWidth      =   6915
      TabIndex        =   8
      Top             =   420
      Visible         =   0   'False
      Width           =   6915
      Begin VB.TextBox txtFind 
         Height          =   345
         Left            =   1050
         TabIndex        =   9
         Top             =   90
         Width           =   3885
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "查找"
         Height          =   210
         Left            =   585
         TabIndex        =   10
         Top             =   165
         Width           =   420
      End
      Begin VB.Image imgSeach 
         Height          =   360
         Left            =   60
         Top             =   75
         Width           =   360
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsItem 
      Height          =   5055
      Left            =   3255
      TabIndex        =   1
      Top             =   930
      Width           =   6900
      _cx             =   12171
      _cy             =   8916
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
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
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   10
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   280
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
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
      ExplorerBar     =   3
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
      Begin MSComctlLib.ImageList imgSort 
         Left            =   510
         Top             =   255
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   9
         ImageHeight     =   8
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmStuffSelect.frx":058A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmStuffSelect.frx":0A64
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame fraInfo 
      Height          =   480
      Left            =   30
      TabIndex        =   6
      Top             =   -75
      Width           =   10155
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "###"
         Height          =   210
         Left            =   225
         TabIndex        =   7
         Top             =   180
         Width           =   315
      End
   End
   Begin VB.CheckBox chkSub 
      Caption         =   "显示所有下级项目(&S)"
      Height          =   210
      Left            =   555
      TabIndex        =   4
      Top             =   6210
      Width           =   2295
   End
   Begin VB.Frame fraLR 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4785
      Left            =   3195
      MousePointer    =   9  'Size W E
      TabIndex        =   5
      Top             =   480
      Width           =   45
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   380
      Left            =   7365
      TabIndex        =   3
      Top             =   6180
      Width           =   1250
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Enabled         =   0   'False
      Height          =   380
      Left            =   6105
      TabIndex        =   2
      Top             =   6180
      Width           =   1250
   End
   Begin MSComctlLib.ImageList img16 
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
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffSelect.frx":0F3E
            Key             =   "Close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffSelect.frx":14D8
            Key             =   "Expend"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffSelect.frx":1A72
            Key             =   "成药"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffSelect.frx":200C
            Key             =   "诊疗"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffSelect.frx":25A6
            Key             =   "草药"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffSelect.frx":2B40
            Key             =   "方案"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvw_s 
      Height          =   5565
      Left            =   15
      TabIndex        =   0
      Top             =   435
      Width           =   3180
      _ExtentX        =   5609
      _ExtentY        =   9816
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "img16"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Shape Shp 
      Height          =   405
      Left            =   3195
      Top             =   6105
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   1
      X1              =   -15
      X2              =   11985
      Y1              =   6045
      Y2              =   6045
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      Index           =   0
      X1              =   0
      X2              =   12000
      Y1              =   6060
      Y2              =   6060
   End
End
Attribute VB_Name = "frmStuffSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String
Private mstrPrivsOpt As String '记帐操作1150模块的授权功能
Private mint险类 As Integer
Private mint病人来源 As Integer
Private mstr输入 As String
Private mlngHwnd As Long
Private mstr特准项目 As String
Private mrsItem As ADODB.Recordset
Private mlng项目ID As Long
Private mblnOK As Boolean
Private mstrLike As String
Private mstrSaveTag As String
Private mstrPreNode As String
Private mlng虚拟库房ID As Long
Private mblnClick As Boolean
Private mrsSel As ADODB.Recordset
Private mblnMyStyle As Boolean
Private mint简码 As Integer
Private mbln包含核算材料 As Boolean
Public Function ShowSelect(frmParent As Object, ByVal strPrivs As String, _
    ByVal int病人来源 As Integer, ByVal int险类 As Integer, _
    Optional ByVal str输入 As String, _
    Optional ByVal lngHwnd As Long, _
    Optional ByVal str特准项目 As String, _
    Optional ByVal lng虚拟库房ID As Long, _
    Optional ByVal bln包含核算材料 As Boolean = False, _
    Optional rsReturnSel As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示卫生材料项目选择器
    '入参:int病人来源=指病人来源,1-门诊,2-住院
    '     str输入=输入匹配的内容,如果没有则为选择器方式,否则为列表方式
    '     lngHwnd=用于列表定位的输入框的句柄
    '     str特准项目=用于医保病人
    '     bln包含核算材料-是否包含卫材料中的核算材料
    '出参:rsReturnSel-返回被选中的数据(虚拟库房ID,收费项目ID,编码,名称,规格,批次,商品条码,内部条码,可用库存,)
    '返回:如果没有数据(已提示),或取消,则返回true；否则False
    '编制:刘兴洪
    '日期:2010-12-13 10:03:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, strKind As String
    mstrPrivs = strPrivs: mbln包含核算材料 = bln包含核算材料
    mstrPrivsOpt = GetInsidePrivs(Enum_Inside_Program.p住院记帐操作)
    mint病人来源 = int病人来源: mint险类 = int险类: mstr输入 = str输入
    mlngHwnd = lngHwnd: mstr特准项目 = str特准项目: mblnOK = False
    mstrSaveTag = IIF(mstr输入 <> "", 1, 0)
    mlng虚拟库房ID = lng虚拟库房ID: Set mrsSel = Nothing
    mblnMyStyle = Val(zlDatabase.GetPara("使用个性化风格")) = 1
    mint简码 = Val(zlDatabase.GetPara("简码方式")) '简码匹配方式：0-拼音,1-五笔

    On Error Resume Next
    Me.Show 1, frmParent
    On Error GoTo 0
    Set rsReturnSel = mrsSel
    ShowSelect = mblnOK

End Function

Private Sub chkSub_Click()
    If Not Visible Then Exit Sub
    vsItem.SetFocus
    Call FillList(True)
End Sub

Private Sub cmdCancel_Click()
    mblnOK = False:    Unload Me
End Sub

Private Sub cmdOK_Click()
    If BulidingRec = False Then
        Set mrsSel = Nothing
        Exit Sub
    End If
    mblnOK = True: Unload Me
End Sub
Private Function BulidingRec() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取选中的数据,以记录集的形式反馈
    '编制:刘兴洪
    '日期:2010-12-14 09:33:42
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, lngID As Long
    Set mrsSel = New ADODB.Recordset
    With vsItem
        lngID = Val(.TextMatrix(.Row, 1))
        If lngID = 0 Then Exit Function
    End With
    With mrsSel
        If .State = 1 Then .Close
        .Fields.Append "虚拟库房ID", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "收费项目ID", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "编码", adLongVarChar, 30, adFldIsNullable
        .Fields.Append "名称", adLongVarChar, 200, adFldIsNullable
        .Fields.Append "规格", adLongVarChar, 100, adFldIsNullable
        .Fields.Append "批次", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "商品条码", adLongVarChar, 50, adFldIsNullable
        .Fields.Append "内部条码", adLongVarChar, 20, adFldIsNullable
        .Fields.Append "可用库存", adDouble, , adFldIsNullable
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    With vsItem
        mrsSel.AddNew
        mrsSel!虚拟库房id = mlng虚拟库房ID
        mrsSel!收费项目ID = lngID
        mrsSel!编码 = Trim(.TextMatrix(.Row, .ColIndex("编码")))
        mrsSel!名称 = Trim(.TextMatrix(.Row, .ColIndex("名称")))
        mrsSel!规格 = Trim(.TextMatrix(.Row, .ColIndex("规格")))
        mrsSel!批次 = Val(.TextMatrix(.Row, .ColIndex("批次")))
        mrsSel!商品条码 = Trim(.TextMatrix(.Row, .ColIndex("商品条码")))
        mrsSel!内部条码 = Trim(.TextMatrix(.Row, .ColIndex("内部条码")))
        mrsSel!可用库存 = Val(.TextMatrix(.Row, .ColIndex("可用库存")))
        mrsSel.Update
    End With
    BulidingRec = True
End Function

Private Sub Form_Activate()
    If Not tvw_s.Visible Then vsItem.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngIdx As Long
    If KeyCode = vbKeyEscape Then
        Call cmdCancel_Click
    End If
End Sub

Private Sub Form_Load()
    Dim lngScrW As Long, lngScrH As Long, lngColW As Long
    Dim vRect As RECT, strIDs As String, i As Long
    Dim lngUpH As Long, lngDnH As Long
    
    Call RestoreWinState(Me, App.ProductName, mstrSaveTag)
    
    mblnOK = False: mblnClick = True: mstrPreNode = ""
    mlng项目ID = 0: mstrLike = IIF(Val(zlDatabase.GetPara("输入匹配")) = 0, "%", "")

    If mstr输入 = "" Then
        '读取类别失败,已提示,非取消退出
        If Not FillTree Then
            mblnOK = True: Unload Me: Exit Sub
        End If
        '无类别,提示,非取消退出
        If tvw_s.Nodes.Count = 0 Then
            MsgBox "没有设置相关卫生材料项目类别,请先到卫生材料目录管理中设置。", vbInformation, gstrSysName
            mblnOK = True: Unload Me: Exit Sub
        End If
    Else
        fraInfo.Visible = False: tvw_s.Visible = False
        fraLR.Visible = False: chkSub.Visible = False
        cmdOK.Visible = False: cmdCancel.Visible = False
        Line1(0).Visible = False: Line1(1).Visible = False: Shp.Visible = True

        '填充匹配数据
        Call FillList(True, strIDs)
        If mrsItem Is Nothing Then
            Unload Me: Exit Sub
        ElseIf mrsItem.RecordCount = 1 Then
            '只有一个项目时,直接返回
            cmdOK_Click
            mblnOK = True: Exit Sub
        ElseIf mrsItem.RecordCount > 0 Then
            '多行是同一个项目时,直接返回
            If mstr输入 <> "" Then
                If UBound(Split(strIDs, ",")) = 1 Then
                    mlng项目ID = Val(vsItem.TextMatrix(vsItem.Row, 1))
                    Call cmdOK_Click
                     Exit Sub
                End If
            End If
            
            vsItem.Appearance = flexFlat
            Call FormSetCaption(Me, False, False)
            Call GetWindowRect(mlngHwnd, vRect) '输入框位置
            vRect.Left = vRect.Left - 2: vRect.Top = vRect.Top - 4
            vRect.Bottom = vRect.Bottom + 4
            
            '设置窗体尺寸和位置
            '计算宽度
            Me.Left = vRect.Left * Screen.TwipsPerPixelX
            lngScrW = GetSystemMetrics(SM_CXVSCROLL) * Screen.TwipsPerPixelX + 60 '+3D边框
            For i = 0 To vsItem.Cols - 1
                lngColW = lngColW + IIF(vsItem.ColHidden(i), 0, vsItem.ColWidth(i))
            Next
            If Me.Left + lngColW + lngScrW > Screen.Width - lngScrW Then
                Me.Width = Screen.Width - lngScrW - Me.Left
            Else
                Me.Width = lngColW + lngScrW
            End If
            
            '计算高度
            lngScrH = GetSystemMetrics(SM_CYFULLSCREEN) * Screen.TwipsPerPixelY '屏幕可用高度
            lngUpH = vRect.Top * Screen.TwipsPerPixelY '上面可用高度
            lngDnH = lngScrH - vRect.Bottom * Screen.TwipsPerPixelY '下面可用高度
            Me.Height = vsItem.Rows * vsItem.RowHeight(0) + 375 '+类别卡片高度
            If Me.Height < 1500 Then Me.Height = 1500 '窗体最小高度
            If Me.Height > lngUpH And Me.Height > lngDnH Then
                Me.Height = IIF(lngUpH < lngDnH, lngDnH, lngUpH)
            End If
            If Me.Height > lngScrH / 2 Then Me.Height = lngScrH / 2 '窗体最大高度
            If Me.Height <= lngDnH Then
                Me.Top = vRect.Bottom * Screen.TwipsPerPixelY
            ElseIf Me.Height <= lngUpH Then
                Me.Top = vRect.Top * Screen.TwipsPerPixelY - Me.Height
            End If
            
            Call Form_Resize
        Else
            '无数据,提示,非取消退出
            MsgBox "没有找到与输入相符的卫生材料项目。", vbInformation, gstrSysName
            mblnOK = True: Unload Me: Exit Sub
        End If
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If mstr输入 = "" Then
        fraInfo.Left = 0
        fraInfo.Width = Me.ScaleWidth
        
        tvw_s.Left = 0
        tvw_s.Top = fraInfo.Top + fraInfo.Height + 15
        tvw_s.Height = Me.ScaleHeight - tvw_s.Top - 690
        
        fraLR.Top = tvw_s.Top
        fraLR.Left = tvw_s.Left + tvw_s.Width
        fraLR.Height = tvw_s.Height
        
        With picStuffItem
            .Top = tvw_s.Top: .Left = fraLR.Left + fraLR.Width: .Width = Me.ScaleWidth - tvw_s.Width - fraLR.Width
        End With
        vsItem.Top = IIF(picStuffItem.Visible, picStuffItem.Top + picStuffItem.Height, tvw_s.Top)
        vsItem.Left = fraLR.Left + fraLR.Width
        vsItem.Width = Me.ScaleWidth - tvw_s.Width - fraLR.Width
        vsItem.Height = tvw_s.Height - IIF(picStuffItem.Visible, picStuffItem.Height, 0)
 
        Line1(0).X1 = 0: Line1(0).X2 = Me.ScaleWidth
        Line1(0).Y1 = tvw_s.Top + tvw_s.Height + 60: Line1(0).Y2 = Line1(0).Y1
        
        Line1(1).X1 = Line1(0).X1: Line1(1).X2 = Line1(0).X2
        Line1(1).Y1 = Line1(0).Y1 - 15: Line1(1).Y2 = Line1(1).Y1
        
        cmdOK.Top = Line1(1).Y1 + 135
        cmdCancel.Top = cmdOK.Top
        
        If Me.ScaleWidth - cmdCancel.Width * 1.5 < 4100 Then
            cmdCancel.Left = 4100
        Else
            cmdCancel.Left = Me.ScaleWidth - cmdCancel.Width * 1.5
        End If
        cmdOK.Left = cmdCancel.Left - cmdOK.Width
        
        chkSub.Top = cmdOK.Top + (cmdOK.Height - chkSub.Height) / 2
    Else
        Shp.Left = 0
        Shp.Top = 0
        Shp.Width = Me.ScaleWidth
        Shp.Height = Me.ScaleHeight
        With picStuffItem
            .Top = 0: .Left = 0: .Width = ScaleWidth
        End With
        
        vsItem.Left = 0
        vsItem.Top = IIF(picStuffItem.Visible, picStuffItem.Top + picStuffItem.Height, 0)
        vsItem.Width = Me.ScaleWidth
        vsItem.Height = Me.ScaleHeight - IIF(picStuffItem.Visible, picStuffItem.Height, 0)
    End If
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mrsItem = Nothing
    Call SaveColPosition
    Call SaveColWidth
    Call SaveWinState(Me, App.ProductName, mstrSaveTag)
End Sub

Private Sub fraLR_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If tvw_s.Width + X < 1000 Or vsItem.Width - X < 1000 Then Exit Sub
        fraLR.Left = fraLR.Left + X
        tvw_s.Width = tvw_s.Width + X
        vsItem.Left = vsItem.Left + X
        vsItem.Width = vsItem.Width - X
        picStuffItem.Left = vsItem.Left
        picStuffItem.Width = vsItem.Width
        Me.Refresh
    End If
End Sub

Private Function FillTree() As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim objNode As Node, strTmp As String
    Dim str类型 As String
 
     
    '卫材料按诊疗分类
     strSQL = _
        "   Select 0 as 级,类型,To_Number('99999999'||类型) as ID,-NULL as 上级ID, '卫生材料' as 名称" & _
        "   From 诊疗分类目录  " & _
        "   Where 类型=7 And (撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null)" & _
        "   Group by 类型" & _
        "   Union ALL " & _
        "   Select Level as 级,类型,-ID as ID," & _
        "       Nvl(-上级ID,To_Number('99999999'||类型)) as 上级ID,'['||编码||']'||名称 as 名称" & _
        "   From 诊疗分类目录 " & _
        "   Where  类型=7   And (撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or 撤档时间 Is Null)" & _
        "   Start With 上级ID is NULL Connect by Prior ID=上级ID"
                
     strSQL = strSQL & " Order by 级,类型,名称"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name)
    
    For i = 1 To rsTmp.RecordCount
        If IsNull(rsTmp!上级ID) Then
            Set objNode = tvw_s.Nodes.Add(, , "_" & rsTmp!ID, rsTmp!名称, "Close")
        Else
            Set objNode = tvw_s.Nodes.Add("_" & rsTmp!上级ID, 4, "_" & rsTmp!ID, rsTmp!名称, "Close")
        End If
        objNode.Tag = rsTmp!类型 '存放分类类型: 7-卫生材料
        objNode.ExpandedImage = "Expend"
        rsTmp.MoveNext
    Next
    If tvw_s.Nodes.Count > 0 Then
        tvw_s.Nodes(1).Expanded = True
        If tvw_s.Nodes(1).Children > 0 Then
            tvw_s.Nodes(1).Child.Selected = True
        Else
            tvw_s.Nodes(1).Selected = True
        End If
        tvw_s.SelectedItem.EnsureVisible
        Call tvw_s_NodeClick(tvw_s.SelectedItem)
    End If
    FillTree = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub vsItem_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow >= vsItem.FixedRows Then
        cmdOK.Enabled = Val(vsItem.TextMatrix(NewRow, 1)) <> 0
    Else
        cmdOK.Enabled = False
    End If
End Sub

Private Sub vsItem_AfterSort(ByVal Col As Long, Order As Integer)
    Dim strType As String, i As Long
    
    With vsItem
        .Cell(flexcpPicture, 0, 0, 0, .Cols - 1) = Nothing
        If Order Mod 2 = 1 Then
            .Cell(flexcpPicture, 0, Col) = imgSort.ListImages(1).Picture
        Else
            .Cell(flexcpPicture, 0, Col) = imgSort.ListImages(2).Picture
        End If
        
        If Val(.TextMatrix(.Row, 1)) <> 0 Then
            .Redraw = flexRDNone
            For i = 1 To .Rows - 1
                .TextMatrix(i, 0) = i
            Next
            .Redraw = flexRDDirect
            Call vsItem_AfterRowColChange(-1, -1, .Row, .Col)
        End If
            
        '因为可能列顺序改变,所以保存原始列号
        If mstr输入 = "" Then strType = tvw_s.SelectedItem.Tag
        SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & mstrSaveTag & "\VSFlexGrid", .Name & strType & "ColSort", .ColData(Col) & "," & Order
    End With
End Sub

Private Sub vsItem_BeforeSort(ByVal Col As Long, Order As Integer)
    '强制编码列按字符串排序
    If vsItem.TextMatrix(0, Col) = "编码" Then
        If Order = 1 Then Order = 7
        If Order = 2 Then Order = 8
    End If
End Sub

Private Sub vsItem_DblClick()
    If vsItem.MouseRow >= vsItem.FixedRows Then
        Call vsItem_KeyPress(13)
    End If
End Sub

Private Sub vsItem_KeyPress(KeyAscii As Integer)
    Static strIdx As String
    Static sngTim As Single
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If cmdOK.Enabled Then cmdOK_Click
    Else
        If KeyAscii >= 48 And KeyAscii <= 57 Then
            If Abs(Timer - sngTim) > 0.5 Then
                strIdx = ""
            End If
            sngTim = Timer
            strIdx = strIdx & Chr(KeyAscii)
            KeyAscii = 0
            
            If Len(strIdx) > 4 Then strIdx = Left(strIdx, 4)
            
            If vsItem.Rows - 1 >= CInt(strIdx) And CInt(strIdx) > 0 Then
                vsItem.Row = Val(strIdx)
                vsItem.ShowCell vsItem.Row, vsItem.Col
            End If
        End If
    End If
End Sub
 

Private Sub tvw_s_NodeClick(ByVal Node As MSComctlLib.Node)
    If Node.Key = mstrPreNode Then Exit Sub
    '结点改变时,保存当前顺序(分类型)
    If Visible Then
        Call SaveColPosition(tvw_s.Nodes(mstrPreNode).Tag)
        Call SaveColWidth(tvw_s.Nodes(mstrPreNode).Tag)
    End If
    mstrPreNode = Node.Key
        
    Call FillList(True)
End Sub

Private Function GetTreePath(ByVal objNode As Node) As String
'功能：获取结点的路径串
    Dim tmpNode As Node, strTmp As String
    Set tmpNode = objNode
    Do While Not tmpNode Is Nothing
        strTmp = NeedName(Replace(tmpNode.Text, Chr(13), "")) & "\" & strTmp
        Set tmpNode = tmpNode.Parent
    Loop
    GetTreePath = strTmp
End Function


Private Sub SaveColPosition(Optional ByVal strType As String)
'功能：保存列顺序:列号,顺序|...
'说明：应放在SaveWinState之前,以在不使用个性化时从注册表清除
    Dim strPos As String, i As Long
        
    If Not mblnMyStyle Then Exit Sub
    
    With vsItem
        For i = 0 To .Cols - 1
            strPos = strPos & "|" & .ColData(i) & "," & i
        Next
        
        If mstr输入 = "" And strType = "" And Not tvw_s.SelectedItem Is Nothing Then strType = tvw_s.SelectedItem.Tag
        SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & mstrSaveTag & "\VSFlexGrid", .Name & strType & "ColPosition", Mid(strPos, 2)
    End With
End Sub

Private Sub SaveColWidth(Optional ByVal strType As String)
'功能：保存列宽度
'说明：应放在SaveWinState之前,以在不使用个性化时从注册表清除
    Dim strPos As String, i As Long
        
    If Not mblnMyStyle Then Exit Sub
    If mstr输入 = "" And strType = "" And Not tvw_s.SelectedItem Is Nothing Then strType = tvw_s.SelectedItem.Tag
    Call SaveFlexState(vsItem, App.ProductName & Me.Name & strType)
End Sub

Private Sub RestoreColWidth()
'功能：恢复列宽度
'说明：应放在恢复列序之后
    Dim strType As String
    
    If Not mblnMyStyle Then Exit Sub
    
    If mstr输入 = "" Then strType = tvw_s.SelectedItem.Tag
    Call RestoreFlexState(vsItem, App.ProductName & Me.Name & strType)
End Sub


Private Sub RestoreColPosition()
'功能：恢复列顺序
'说明：应放在排序处理之前
    Dim rsPos As New ADODB.Recordset
    Dim strType As String, strPos As String
    Dim i As Long, j As Long
    
    If Not mblnMyStyle Then Exit Sub
    
    With vsItem
        If mstr输入 = "" Then strType = tvw_s.SelectedItem.Tag
        strPos = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & mstrSaveTag & "\VSFlexGrid", .Name & strType & "ColPosition", "")
        If strPos <> "" Then
            rsPos.Fields.Append "Col", adBigInt
            rsPos.Fields.Append "Position", adBigInt
            rsPos.CursorLocation = adUseClient
            rsPos.LockType = adLockOptimistic
            rsPos.CursorType = adOpenStatic
            rsPos.Open
            
            For i = 0 To UBound(Split(strPos, "|"))
                rsPos.AddNew
                rsPos!Col = Split(Split(strPos, "|")(i), ",")(0)
                rsPos!Position = Split(Split(strPos, "|")(i), ",")(1)
                rsPos.Update
            Next
            rsPos.Sort = "Position"
            
            'ColPosition:>=0,ReadOnly,改变后相关列号也改变
            For i = 1 To rsPos.RecordCount
                For j = i - 1 To .Cols - 1
                    If .ColData(j) = rsPos!Col Then Exit For
                Next
                If j <= .Cols - 1 Then
                    .ColPosition(j) = rsPos!Position
                End If
                rsPos.MoveNext
            Next
        End If
    End With
End Sub

Private Sub RestoreColSort()
'功能：排序处理
    Dim strType As String, strSort As String, i As Long
        
    With vsItem
        Set .Cell(flexcpPicture, 0, 0, 0, .Cols - 1) = Nothing
        .Cell(flexcpPictureAlignment, 0, 0, 0, .Cols - 1) = 7
        If mblnMyStyle Then
            If mstr输入 = "" Then strType = tvw_s.SelectedItem.Tag
            strSort = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & mstrSaveTag & "\VSFlexGrid", .Name & strType & "ColSort", "")
            If strSort <> "" Then
                '因为可能调整列顺序,所以查找真实的排序列
                For i = 0 To .Cols - 1
                    If .ColData(i) = Val(Split(strSort, ",")(0)) Then Exit For
                Next
                If i <= .Cols - 1 Then
                    .Col = i
                    .Sort = Val(Split(strSort, ",")(1))
                    
                    If Val(Split(strSort, ",")(1)) Mod 2 = 1 Then
                        .Cell(flexcpPicture, 0, i) = imgSort.ListImages(1).Picture
                    Else
                        .Cell(flexcpPicture, 0, i) = imgSort.ListImages(2).Picture
                    End If
                End If
            End If
        End If
    End With
End Sub

Private Function FillList(Optional ByVal blnClass As Boolean, Optional strIDs As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据当前界面条件装入诊疗项目目录
    '入参:blnClass=是否重建分类卡(应在树形项目改变时才重建)
    '       strIDs=读取的项目ID集,用于判断输入时是否别名不同的同一个收费项目
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2010-12-13 10:26:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objTab As MSComctlLib.Tab
    Dim objNode As Node, objItem As ListItem
    Dim arrClass As Variant, strClass As String
    Dim strInput As String, blnLoad As Boolean
    Dim str分类ID As String, lng药房ID As Long, strStock As String
    Dim strMain As String, strSQL As String
    Dim blnStock As Boolean, strTmp As String, strSQLItem As String
    Dim i As Long, j As Long
    
    Dim int类型 As Integer, lng诊疗分类ID As Long, lng收费分类ID As Long
    
    strIDs = ""
    Set objNode = tvw_s.SelectedItem '输入匹配时,为Nothing
    
    '清除项目清单及分类卡片
    '------------------------------------------------------------------------
    vsItem.Rows = vsItem.FixedRows
    vsItem.Rows = vsItem.FixedRows + 1
 
    Me.Refresh
    
    '共公条件及字段设置
    '------------------------------------------------------------------------
    If mstr输入 = "" Then
        int类型 = Val(objNode.Tag)
        lng诊疗分类ID = -1 * Val(Mid(objNode.Key, 2))
        lng收费分类ID = Val(Mid(objNode.Key, 2))
        
        '树形中的分类ID
        If chkSub.Value = 1 Then
            '显示下级的项目
            If Val(objNode.Tag) > 0 Then
                '诊疗分类目录
                If Mid(objNode.Key, 2) = "99999999" & objNode.Tag Then
                    str分类ID = " And E.分类ID IN(Select ID From 诊疗分类目录 Where 类型=[1])"
                Else
                    str分类ID = " And E.分类ID IN(Select ID From 诊疗分类目录 Start With ID=[2] Connect by Prior ID=上级ID)"
                End If
            Else
                '收费分类目录
                str分类ID = " And A.分类ID IN(Select ID From 收费分类目录 Start With ID=[3] Connect by Prior ID=上级ID)"
            End If
        Else
            If Val(objNode.Tag) > 0 Then
                '诊疗分类目录
                str分类ID = " And E.分类ID=[2]"
            Else
                '收费分类目录
                str分类ID = " And A.分类ID=[3]"
            End If
        End If
    Else
        '输入匹配
        If Len(mstr输入) < 2 Then mstrLike = "" '优化
    End If

    
    strSQLItem = " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL) And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)"

            
    '读取数据
    '------------------------------------------------------------------------
    '卫材收费项目部份
    '因为从诊疗中分类,所以单独处理。目前暂不管库存显示


    If mstr输入 <> "" Then
        strInput = " And (A.编码 Like [6] Or B.名称 Like [7] Or B.简码 Like [7] Or M.商品条码=[4] or M.内部条码=[4] ) And B.码类=[8]"
        If IsNumeric(mstr输入) Then                         '10,11.输入全是数字时只匹配编码,内部码,商品码
            If Mid(gstrMatchMode, 1, 1) = "1" Then strInput = " And (A.编码 Like [6] Or M.商品条码=[4] or M.内部条码=[4]) And B.码类=[8]"
        ElseIf zlCommFun.IsCharAlpha(mstr输入) Then         '01,11.输入全是字母时只匹配简码
            If Mid(gstrMatchMode, 2, 1) = "1" Then strInput = " And B.简码 Like [7] And B.码类=[8]"
        ElseIf zlCommFun.IsCharChinese(mstr输入) Then
            strInput = " And B.名称 Like [7] And B.码类=[8]"
        End If
        
        If mint险类 = 0 Then
            strMain = _
            " Select Distinct A.ID,A.类别,A.编码,B.名称,B.简码," & _
            "       A.计算单位,A.规格,A.产地,A.费用类型,Null 医保大类,A.说明,A.是否变价, " & _
            "       M.批次,M.商品条码,M.内部条码,nvl(M.可用数量,0) as 可用库存 " & _
            " From 收费项目别名 B,收费项目目录 A,药品库存 M " & _
            " Where A.服务对象 IN([9],3) And A.ID=B.收费细目ID And A.类别='4'" & mstr特准项目 & strInput & strSQLItem & _
            "           And A.ID=M.药品ID And M.库房ID=[11]   And (  M.效期 Is Null Or M.效期 > Trunc(Sysdate)) " & _
            "           And nvl(M.可用数量,0) >0 "
        Else
            strMain = _
            " Select Distinct A.ID,A.类别,A.编码,B.名称,B.简码," & _
            "       A.计算单位,A.规格,A.产地,A.费用类型,D.名称 医保大类,A.说明,A.是否变价," & _
            "       M.批次,M.商品条码,M.内部条码,nvl(M.可用数量,0) as 可用库存 " & _
            " From 收费项目别名 B,收费项目目录 A,保险支付项目 C,保险支付大类 D,药品库存 M" & _
            " Where A.服务对象 IN([9],3) And A.ID=B.收费细目ID And A.类别='4'" & mstr特准项目 & strInput & strSQLItem & _
            "           And A.ID=C.收费细目ID(+) And C.险类(+)=[12] And C.大类ID=D.ID(+)" & _
            "           And A.ID=M.药品ID And M.库房ID=[11]   And (  M.效期 Is Null Or M.效期 > Trunc(Sysdate)) " & _
            "           And nvl(M.可用数量,0) >0 "
        End If
    Else
        If mint险类 = 0 Then
            strMain = _
            "   Select A.ID,A.类别,A.编码,A.名称,A.计算单位,A.规格,A.产地,A.费用类型,Null 医保大类,A.说明,A.是否变价," & _
            "       M.批次,M.商品条码,M.内部条码,nvl(M.可用数量,0) as 可用库存 " & _
            "   From 收费项目目录 A,药品库存 M" & _
            "   Where   A.服务对象 IN([9],3) And A.类别='4'" & mstr特准项目 & strSQLItem & _
            "           And A.ID=M.药品ID And M.库房ID=[11]   And (  M.效期 Is Null Or M.效期 > Trunc(Sysdate)) " & _
            "           And nvl(M.可用数量,0) >0 "
        Else
            strMain = _
            "   Select A.ID,A.类别,A.编码,A.名称,A.计算单位,A.规格,A.产地,A.费用类型,D.名称 医保大类,A.说明,A.是否变价," & _
            "       M.批次,M.商品条码,M.内部条码,nvl(M.可用数量,0) as 可用库存 " & _
            "   From 收费项目目录 A,保险支付项目 C,保险支付大类 D,药品库存 M" & _
            "   Where A.服务对象 IN([9],3) And A.类别='4'" & mstr特准项目 & strSQLItem & _
            "       And A.ID=C.收费细目ID(+) And C.险类(+)=[12] And C.大类ID=D.ID(+)" & _
            "           And A.ID=M.药品ID And M.库房ID=[11]   And (  M.效期 Is Null Or M.效期 > Trunc(Sysdate)) " & _
            "           And nvl(M.可用数量,0) >0 "
        End If
    End If
        
        '卫材不受此系统参数和权限影响:分离下药时指定药房时限定库存,权限:不检查库存
    strSQL = "" & _
    "   Select A.ID,A.类别 as 类别ID,B.序号 as 顺序ID,B.名称 as 类别," & _
    "           A.编码,A.名称,NULL as 商品名," & IIF(mstr输入 <> "", "A.简码,", "") & _
    "           A.计算单位 as 单位,A.规格,A.产地,A.费用类型,A.医保大类,A.说明," & _
    "           A.批次,A.商品条码,A.内部条码," & _
    "           Decode(A.是否变价,1,'时价',LTrim(To_Char(Sum(C.现价),'999999" & gstrDecPrice & "'))) as 单价,A.可用库存," & _
                IIF(InStr(1, mstrPrivsOpt, "显示库存") > 0, " To_Char(A.可用库存,'9999990.00000')", "Decode(Sign(A.可用库存),1,'有','无')") & " as 库存" & _
    " From (" & strMain & ") A,收费项目类别 B,收费价目 C,材料特性 D,诊疗项目目录 E" & _
    " Where A.类别=B.编码 And A.ID=C.收费细目ID   " & IIF(mbln包含核算材料, "", " And nvl(D.核算材料 ,0)=0 ") & _
    "           And A.ID=D.材料ID and nvl(D.跟踪在用,0)=1 And D.诊疗ID=E.ID" & str分类ID & _
    "           And Sysdate Between C.执行日期 and Nvl(C.终止日期,To_Date('3000-01-01','YYYY-MM-DD'))" & _
    " Group by A.ID,A.类别,B.序号,B.名称,A.编码,A.名称," & IIF(mstr输入 <> "", "A.简码,", "") & _
    "               A.规格,A.产地,A.费用类型,A.医保大类,A.说明,A.批次,A.商品条码,A.内部条码,A.是否变价,A.计算单位,A.可用库存, To_Char(A.可用库存,'9999990.00000')"
    strSQL = strSQL & " Order by 顺序ID,编码,批次,内部条码,商品条码" '使用类别序号排序
    
    On Error GoTo errH
    Screen.MousePointer = 11
    Set mrsItem = zlDatabase.OpenSQLRecord(strSQL, Me.Name, int类型, lng诊疗分类ID, lng收费分类ID, UCase(mstr输入), "4", _
        UCase(mstr输入) & "%", mstrLike & UCase(mstr输入) & "%", mint简码 + 1, mint病人来源, IIF(gbyt药品名称显示 = 1, 3, 1), _
         mlng虚拟库房ID, mint险类)
    
    '绑定数据
    '--------------------------------------------------------------------------
    vsItem.Redraw = flexRDNone
    vsItem.ScrollBars = flexScrollBarNone
    Set vsItem.DataSource = mrsItem
    vsItem.ScrollBars = flexScrollBarBoth
    If Err.Number = 0 And gcnOracle.Errors.Count > 0 Then
        gcnOracle.Errors.Clear
    End If
    If vsItem.Rows = vsItem.FixedRows Then
        vsItem.Rows = vsItem.FixedRows + 1
    End If
    
    '列属性调整
    vsItem.ColAlignment(0) = 4
    vsItem.Cell(flexcpAlignment, 0, 0, 0, vsItem.Cols - 1) = 4
    vsItem.RowHeight(0) = vsItem.RowHeightMin
    For i = 1 To vsItem.Cols - 1
        vsItem.ColKey(i) = vsItem.TextMatrix(0, i)
        
        If InStr("单价,库存", vsItem.TextMatrix(0, i)) > 0 Then
            vsItem.ColAlignment(i) = 7
        Else
            vsItem.ColAlignment(i) = 1
        End If
        If vsItem.TextMatrix(0, i) Like "*ID" Or vsItem.TextMatrix(0, i) = "可用库存" _
            Or vsItem.ColKey(i) = "商品名" Or vsItem.ColKey(i) = "类别" Then
            vsItem.ColHidden(i) = True
            vsItem.ColWidth(i) = 0
        ElseIf vsItem.ColWidth(i) > 2800 Then
            vsItem.ColWidth(i) = 2800
        ElseIf mrsItem.RecordCount = 0 Then
            vsItem.ColWidth(i) = 1000
        End If
        vsItem.ColData(i) = i '记录原始列号,用于处理列顺序
    Next
    
    '恢复列顺序:应放在排序处理之前
    Call RestoreColPosition
    Call RestoreColWidth
    '排序处理:先排序,以便后面处理行号
    Call RestoreColSort
    
    '卡片相关数据计算
    '------------------------------------------------------------------------
    For i = 1 To mrsItem.RecordCount
        vsItem.TextMatrix(i, 0) = i
        vsItem.RowHeight(i) = vsItem.RowHeightMin
        
        '收集类别卡片信息
        If InStr(strClass & ",", "," & mrsItem!类别ID & mrsItem!类别 & ",") = 0 Then
            strClass = strClass & "," & mrsItem!类别ID & mrsItem!类别
        End If
        
        '收集项目ID:只收集最多2个
        If mstr输入 <> "" Then
            If UBound(Split(strIDs, ",")) < 2 Then
                If InStr(strIDs & ",", "," & mrsItem!ID & ",") = 0 Then
                    strIDs = strIDs & "," & mrsItem!ID
                End If
            End If
        End If
        mrsItem.MoveNext
    Next
 
    
    '根据情况隐藏一些不必要的列
    '刘兴洪 问题:27068和27990 日期:2009-12-24 15:41:04
'    .byt输入药品显示 = Val(zlDatabase.GetPara("输入药品显示", glngSys, 0, "0")) '0-按输入匹配显示，1-固定显示通用名和商品名
'    .byt药品名称显示 = Val(zlDatabase.GetPara("药品名称显示", glngSys, 0, "0")) '：0-显示通用名，1-显示商品名，2-同时显示通用名和商品名
        
    For i = 1 To vsItem.Cols - 1
        If vsItem.TextMatrix(0, i) = "商品名" Then
            If (gbyt输入药品显示 = 0 And mstr输入 <> "") Or (mstr输入 = "" And gbyt药品名称显示 <> 2) Then
                vsItem.ColHidden(i) = True '输入时才显示，选择器直接根据参数来
            End If
        ElseIf vsItem.TextMatrix(0, i) = "简码" Then
            If gbyt输入药品显示 = 1 And mstr输入 <> "" Then
                vsItem.ColHidden(i) = True
            End If
        End If
    Next

    '行号列宽度
    vsItem.ColWidth(0) = Me.TextWidth(vsItem.TextMatrix(vsItem.Rows - 1, 0) & " ")
    If vsItem.ColWidth(0) < 380 Then vsItem.ColWidth(0) = 380
    
    vsItem.Row = vsItem.FixedRows: vsItem.Col = vsItem.FixedCols
    Call vsItem_AfterRowColChange(-1, -1, vsItem.Row, vsItem.Col)
    vsItem.Redraw = flexRDDirect
        
    Call Form_Resize
    
    If mrsItem.RecordCount > 0 Then mrsItem.MoveFirst
    lblInfo.Caption = "当前选择：" & GetTreePath(tvw_s.SelectedItem) & "，共有 " & mrsItem.RecordCount & " 个项目"
        
    Screen.MousePointer = 0
    Exit Function
errH:
    LockWindowUpdate 0
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    cmdOK.Enabled = False
End Function
