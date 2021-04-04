VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmItemSelect 
   AutoRedraw      =   -1  'True
   Caption         =   "收费项目选择器"
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
   Icon            =   "frmItemSelect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   10185
   Begin VSFlex8Ctl.VSFlexGrid vsItem 
      Height          =   5190
      Left            =   3255
      TabIndex        =   1
      Top             =   435
      Width           =   6900
      _cx             =   12171
      _cy             =   9155
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
      FormatString    =   $"frmItemSelect.frx":058A
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
         Left            =   810
         Top             =   810
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
               Picture         =   "frmItemSelect.frx":0617
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmItemSelect.frx":0AF1
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame fraInfo 
      Height          =   480
      Left            =   30
      TabIndex        =   7
      Top             =   -75
      Width           =   10155
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "###"
         Height          =   210
         Left            =   225
         TabIndex        =   8
         Top             =   180
         Width           =   315
      End
   End
   Begin VB.CheckBox chkSub 
      Caption         =   "显示所有下级项目(&S)"
      Height          =   210
      Left            =   555
      TabIndex        =   5
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
      TabIndex        =   6
      Top             =   480
      Width           =   45
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   380
      Left            =   7365
      TabIndex        =   4
      Top             =   6180
      Width           =   1250
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Enabled         =   0   'False
      Height          =   380
      Left            =   6105
      TabIndex        =   3
      Top             =   6180
      Width           =   1250
   End
   Begin MSComctlLib.TabStrip tabClass 
      Height          =   540
      Left            =   3315
      TabIndex        =   2
      Top             =   5460
      Width           =   5355
      _ExtentX        =   9446
      _ExtentY        =   953
      TabWidthStyle   =   2
      TabFixedWidth   =   1764
      TabFixedHeight  =   616
      HotTracking     =   -1  'True
      Placement       =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "全部(0)"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "西成药"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "中成药"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "中草药"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "治疗"
            ImageVarType    =   2
         EndProperty
      EndProperty
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
   Begin MSComctlLib.ImageList img16 
      Left            =   1110
      Top             =   2235
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
            Picture         =   "frmItemSelect.frx":0FCB
            Key             =   "Close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItemSelect.frx":1565
            Key             =   "Expend"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItemSelect.frx":1AFF
            Key             =   "成药"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItemSelect.frx":2099
            Key             =   "诊疗"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItemSelect.frx":2633
            Key             =   "草药"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmItemSelect.frx":2BCD
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
      X1              =   0
      X2              =   12000
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
Attribute VB_Name = "frmItemSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstrPrivs As String
Private mint病人来源 As Integer
Private mbln药房单位 As Boolean
Private mstr类别 As String
Private mstr输入 As String
Private mlngHwnd As Long
Private mstr特准项目 As String

Private mrsItem As ADODB.Recordset
Private mlng西药房 As Long
Private mlng成药房 As Long
Private mlng中药房 As Long
Private mlng项目ID As Long
Private mblnOK As Boolean
Private mstrLike As String
Private mint简码 As Integer
Private mstrSaveTag As String
Private mstrPreNode As String
Private mblnClick As Boolean

Public Function ShowSelect(frmParent As Object, ByVal strPrivs As String, _
    ByVal int病人来源 As Integer, ByVal bln药房单位 As Boolean, _
    ByVal str类别 As String, Optional ByVal str输入 As String, _
    Optional ByVal lngHwnd As Long, Optional ByVal str特准项目 As String) As Long
'功能：显示收费项目选择器
'参数：int病人来源=指病人来源,1-门诊,2-住院
'      bln药房单位=是否按药房单位显示库存和价格
'      str类别="'5','D','Z'..",表示允许选择或当前确定要输入的类别,为空表示所有类别
'      str输入=输入匹配的内容,如果没有则为选择器方式,否则为列表方式
'      lngHwnd=用于列表定位的输入框的句柄
'      str特准项目=用于医保病人
'返回：如果没有数据(已提示),或取消,则返回0；否则收费项目ID
    mstrPrivs = strPrivs
    mint病人来源 = int病人来源
    mbln药房单位 = bln药房单位
    mstr类别 = str类别
    mstr输入 = str输入
    mlngHwnd = lngHwnd
    mstr特准项目 = str特准项目
    
    mstrSaveTag = IIF(mstr输入 <> "", 1, 0)
    
    On Error Resume Next
    Me.Show 1, frmParent
    On Error GoTo 0
    
    If mblnOK Then
        ShowSelect = mlng项目ID
    End If
End Function

Private Sub chkSub_Click()
    If Not Visible Then Exit Sub
    vsItem.SetFocus
    Call FillList(True)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    mlng项目ID = Val(vsItem.TextMatrix(vsItem.Row, 1))
    mblnOK = True: Unload Me
End Sub

Private Sub Form_Activate()
    If Not tvw_s.Visible Then vsItem.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngIdx As Long
    
    If KeyCode = vbKeyEscape Then
        Call cmdCancel_Click
    ElseIf Shift = vbAltMask Then
        If Between(KeyCode, vbKey0, vbKey9) Then
            lngIdx = KeyCode - vbKey0 + 1
        End If
        If tabClass.SelectedItem.Index <> lngIdx And Between(lngIdx, 1, tabClass.Tabs.Count) Then
            tabClass.Tabs(lngIdx).Selected = True
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim lngScrW As Long, lngScrH As Long, lngColW As Long
    Dim vRect As RECT, strIDs As String, i As Long
    Dim lngUpH As Long, lngDnH As Long
    
    Call RestoreWinState(Me, App.ProductName, mstrSaveTag)
    
    mblnOK = False
    mblnClick = True
    mstrPreNode = ""
    mlng项目ID = 0
    mstrLike = IIF(GetSetting("ZLSOFT", "公共模块\操作", "输入匹配", 0) = 0, "%", "")
    mint简码 = Val(GetSetting("ZLSOFT", "私有全局\" & gstrDBUser, "简码生成", 0)) '简码匹配方式：0-拼音,1-五笔

    mlng西药房 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, IIF(mint病人来源 = 2, "住院", "门诊") & "缺省西药房", 0))
    mlng成药房 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, IIF(mint病人来源 = 2, "住院", "门诊") & "缺省成药房", 0))
    mlng中药房 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName, IIF(mint病人来源 = 2, "住院", "门诊") & "缺省中药房", 0))

    If mstr输入 = "" Then
        '读取类别失败,已提示,非取消退出
        If Not FillTree Then
            mblnOK = True: Unload Me: Exit Sub
        End If
        '无类别,提示,非取消退出
        If tvw_s.Nodes.Count = 0 Then
            MsgBox "没有设置相关收费项目类别,请先到收费项目管理中设置。", vbInformation, gstrSysName
            mblnOK = True: Unload Me: Exit Sub
        End If
    Else
        fraInfo.Visible = False
        tvw_s.Visible = False
        fraLR.Visible = False
        chkSub.Visible = False
        cmdOK.Visible = False
        cmdCancel.Visible = False
        Line1(0).Visible = False
        Line1(1).Visible = False
        Shp.Visible = True

        '填充匹配数据
        Call FillList(True, strIDs)
        If mrsItem Is Nothing Then
            Unload Me: Exit Sub
        ElseIf mrsItem.RecordCount = 1 Then
            '只有一个项目时,直接返回
            mlng项目ID = Val(vsItem.TextMatrix(vsItem.Row, 1))
            mblnOK = True: Unload Me: Exit Sub
        ElseIf mrsItem.RecordCount > 0 Then
            '多行是同一个项目时,直接返回
            If mstr输入 <> "" Then
                If UBound(Split(strIDs, ",")) = 1 Then
                    mlng项目ID = Val(vsItem.TextMatrix(vsItem.Row, 1))
                    mblnOK = True: Unload Me: Exit Sub
                End If
            End If
            
            vsItem.Appearance = flexFlat
            Call FormSetCaption(Me, False, False)
            Call GetWindowRect(mlngHwnd, vRect) '输入框位置
            vRect.Left = vRect.Left - 2
            vRect.Top = vRect.Top - 4
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
            MsgBox "没有找到与输入相符的收费项目。", vbInformation, gstrSysName
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
        
        vsItem.Top = tvw_s.Top
        vsItem.Left = fraLR.Left + fraLR.Width
        vsItem.Width = Me.ScaleWidth - tvw_s.Width - fraLR.Width
        vsItem.Height = tvw_s.Height - IIF(tabClass.Visible, 350, 0)
        
        If tabClass.Visible Then
            tabClass.Top = vsItem.Top + vsItem.Height - tabClass.Height + 380
            tabClass.Left = vsItem.Left + 30
            tabClass.Width = vsItem.Width - 60
        End If
        
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
        
        vsItem.Left = 0
        vsItem.Top = 0
        vsItem.Width = Me.ScaleWidth
        vsItem.Height = Me.ScaleHeight - IIF(tabClass.Tabs.Count > 1, 375, 0)
        
        If tabClass.Tabs.Count > 1 Then
            tabClass.Left = vsItem.Left + 60
            tabClass.Width = vsItem.Width - 120
            tabClass.Top = Me.ScaleHeight - tabClass.Height - 30
        End If
    End If
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mrsItem = Nothing
    Call SaveColPosition
    Call SaveColWidth
    Call SaveWinState(Me, App.ProductName, mstrSaveTag)
End Sub

Private Sub fraLR_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If tvw_s.Width + x < 1000 Or vsItem.Width - x < 1000 Then Exit Sub
        fraLR.Left = fraLR.Left + x
        tvw_s.Width = tvw_s.Width + x
        vsItem.Left = vsItem.Left + x
        vsItem.Width = vsItem.Width - x
        tabClass.Left = tabClass.Left + x
        tabClass.Width = tabClass.Width - x
        Me.Refresh
    End If
End Sub

Private Function FillTree() As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String, i As Long
    Dim objNode As Node, strTmp As String
    Dim str类型 As String
    
    '药品和卫材料按诊疗分类
    If mstr类别 = "" Or InStr(mstr类别, "'5'") > 0 Then strTmp = strTmp & ",1"
    If mstr类别 = "" Or InStr(mstr类别, "'6'") > 0 Then strTmp = strTmp & ",2"
    If mstr类别 = "" Or InStr(mstr类别, "'7'") > 0 Then strTmp = strTmp & ",3"
    If mstr类别 = "" Or InStr(mstr类别, "'4'") > 0 Then strTmp = strTmp & ",7"
    str类型 = Mid(strTmp, 2)
    If str类型 <> "" Then
        strSQL = _
            " Select 0 as 级,类型,To_Number('99999999'||类型) as ID,-NULL as 上级ID," & _
            " CHR(13)||Decode(类型,1,'西成药',2,'中成药',3,'中草药',7,'卫生材料') as 名称" & _
            " From 诊疗分类目录 Where Instr([1],','||类型||',')>0" & _
            " Group by 类型"
        strSQL = strSQL & " Union ALL " & _
            " Select Level as 级,类型,-ID as ID," & _
            " Nvl(-上级ID,To_Number('99999999'||类型)) as 上级ID,'['||编码||']'||名称 as 名称" & _
            " From 诊疗分类目录 Where Instr([1],','||类型||',')>0" & _
            " Start With 上级ID is NULL Connect by Prior ID=上级ID"
    End If
                
    '非药品收费项目
    strTmp = mstr类别
    strTmp = Replace(strTmp, "'5'", "")
    strTmp = Replace(strTmp, "'6'", "")
    strTmp = Replace(strTmp, "'7'", "")
    strTmp = Replace(strTmp, "'4'", "")
    strTmp = Trim(Replace(strTmp, ",", ""))
    If strTmp <> "" Or mstr类别 = "" Then
        strSQL = strSQL & IIF(strSQL <> "", " Union ALL ", "") & _
            " Select Level as 级,0 as 类型,ID,上级ID,'['||编码||']'||名称 as 名称" & _
            " From 收费分类目录 Start With 上级ID is NULL Connect by Prior ID=上级ID"
    End If
    strSQL = strSQL & " Order by 级,类型,名称"
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Name, "," & str类型 & ",")
    For i = 1 To rsTmp.RecordCount
        If IsNull(rsTmp!上级ID) Then
            Set objNode = tvw_s.Nodes.Add(, , "_" & rsTmp!ID, rsTmp!名称, "Close")
        Else
            Set objNode = tvw_s.Nodes.Add("_" & rsTmp!上级ID, 4, "_" & rsTmp!ID, rsTmp!名称, "Close")
        End If
        objNode.Tag = rsTmp!类型 '存放分类类型:0-非药品和卫材,1-西成药,2-中成药,3-中草药,7-卫生材料
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
        'tvw_s.Nodes(1).Selected = True
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

Private Sub tabClass_Click()
    If Not mblnClick Then Exit Sub
    Call FillList
    vsItem.SetFocus
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
        
    If Val(GetSetting("ZLSOFT", "私有全局\" & gstrDBUser, "使用个性化风格", 1)) = 0 Then Exit Sub
    
    With vsItem
        For i = 0 To .Cols - 1
            strPos = strPos & "|" & .ColData(i) & "," & i
        Next
        
        If mstr输入 = "" And strType = "" And Not tvw_s.SelectedItem Is Nothing Then strType = tvw_s.SelectedItem.Tag
        SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & mstrSaveTag & "\VSFlexGrid", .Name & strType & "ColPosition", Mid(strPos, 2)
    End With
End Sub

Private Sub RestoreColPosition()
'功能：恢复列顺序
'说明：应放在排序处理之前
    Dim rsPos As New ADODB.Recordset
    Dim strType As String, strPos As String
    Dim i As Long, j As Long
    
    If Val(GetSetting("ZLSOFT", "私有全局\" & gstrDBUser, "使用个性化风格", 1)) = 0 Then Exit Sub
    
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

Private Sub SaveColWidth(Optional ByVal strType As String)
'功能：保存列宽度
'说明：应放在SaveWinState之前,以在不使用个性化时从注册表清除
    Dim strPos As String, i As Long
        
    If Val(GetSetting("ZLSOFT", "私有全局\" & gstrDBUser, "使用个性化风格", 1)) = 0 Then Exit Sub
    If mstr输入 = "" And strType = "" And Not tvw_s.SelectedItem Is Nothing Then strType = tvw_s.SelectedItem.Tag
    Call SaveFlexState(vsItem, App.ProductName & Me.Name & strType)
End Sub

Private Sub RestoreColWidth()
'功能：恢复列宽度
'说明：应放在恢复列序之后
    Dim strType As String
    
    If Val(GetSetting("ZLSOFT", "私有全局\" & gstrDBUser, "使用个性化风格", 1)) = 0 Then Exit Sub
    
    If mstr输入 = "" Then strType = tvw_s.SelectedItem.Tag
    Call RestoreFlexState(vsItem, App.ProductName & Me.Name & strType)
End Sub

Private Sub RestoreColSort()
'功能：排序处理
    Dim strType As String, strSort As String, i As Long
        
    With vsItem
        Set .Cell(flexcpPicture, 0, 0, 0, .Cols - 1) = Nothing
        .Cell(flexcpPictureAlignment, 0, 0, 0, .Cols - 1) = 7
        If Val(GetSetting("ZLSOFT", "私有全局\" & gstrDBUser, "使用个性化风格", 1)) <> 0 Then
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
'功能：根据当前界面条件装入诊疗项目目录
'参数：blnClass=是否重建分类卡(应在树形项目改变时才重建)
'      strIDs=读取的项目ID集,用于判断输入时是否别名不同的同一个收费项目
    Dim objTab As MSComctlLib.Tab
    Dim objNode As Node, objItem As ListItem
    Dim arrClass As Variant, strClass As String
    Dim strInput As String, blnLoad As Boolean
    Dim str类别 As String, str分类ID As String
    Dim lng药房ID As Long, strStock As String
    Dim str药房单位 As String, str药房包装 As String
    Dim strMain As String, strSQL As String
    Dim blnStock As Boolean, strTmp As String
    Dim i As Long, j As Long
    
    Dim int类型 As Integer, lng诊疗分类ID As Long, lng收费分类ID As Long
    
    strIDs = ""
    Set objNode = tvw_s.SelectedItem '输入匹配时,为Nothing
    
    '清除项目清单及分类卡片
    '------------------------------------------------------------------------
    vsItem.Rows = vsItem.FixedRows
    vsItem.Rows = vsItem.FixedRows + 1
    If blnClass Then
        mblnClick = False
        tabClass.SelectedItem = tabClass.Tabs(1)
        For i = tabClass.Tabs.Count To 2 Step -1
            tabClass.Tabs.Remove i
        Next
        mblnClick = True
    End If
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
        
        '树形中的类型确定类别
        If Val(objNode.Tag) = 1 Then
            str类别 = "5"
        ElseIf Val(objNode.Tag) = 2 Then
            str类别 = "6"
        ElseIf Val(objNode.Tag) = 3 Then
            str类别 = "7"
        ElseIf Val(objNode.Tag) = 7 Then
            str类别 = "4"
        End If
    Else
        '输入匹配
        If Len(mstr输入) < 2 Then mstrLike = "" '优化
    End If
    
    '类别卡片确定类别
    If tabClass.SelectedItem.Key <> "" Then
        str类别 = Mid(tabClass.SelectedItem.Key, 2)
    End If
            
    '读取数据
    '------------------------------------------------------------------------
    '药品收费项目部份
    blnLoad = False
    If str类别 <> "" Then
        blnLoad = InStr(",5,6,7,", str类别) > 0
    ElseIf mstr输入 <> "" Then '选择器中药品必然可以确定类别(分类或卡片)
        If mstr类别 = "" Or InStr(mstr类别, "'5'") > 0 Then blnLoad = True
        If mstr类别 = "" Or InStr(mstr类别, "'6'") > 0 Then blnLoad = True
        If mstr类别 = "" Or InStr(mstr类别, "'7'") > 0 Then blnLoad = True
    End If
    If blnLoad Then
        '药品库存
        blnStock = True
        If mstr输入 = "" Then
            '根据当前选择的分类可以确定药品类别
            If Val(objNode.Tag) = 1 Then
                lng药房ID = mlng西药房
                If lng药房ID = 0 Then blnStock = False
            ElseIf Val(objNode.Tag) = 2 Then
                lng药房ID = mlng成药房
                If lng药房ID = 0 Then blnStock = False
            ElseIf Val(objNode.Tag) = 3 Then
                lng药房ID = mlng中药房
                If lng药房ID = 0 Then blnStock = False
            End If
        Else
            '根据调用程序限制类别
            If mstr类别 = "'5'" Or str类别 = "5" Then
                lng药房ID = mlng西药房
                If lng药房ID = 0 Then blnStock = False
            ElseIf mstr类别 = "'6'" Or str类别 = "6" Then
                lng药房ID = mlng成药房
                If lng药房ID = 0 Then blnStock = False
            ElseIf mstr类别 = "'7'" Or str类别 = "7" Then
                lng药房ID = mlng中药房
                If lng药房ID = 0 Then blnStock = False
            End If
        End If
        If lng药房ID <> 0 Then
            strStock = _
                " Select A.药品ID,Sum(Nvl(A.可用数量,0)) as 库存" & _
                " From 药品库存 A" & _
                " Where A.性质 = 1 And A.库房ID=[11]" & _
                " And (Nvl(A.批次, 0) = 0 Or A.效期 Is Null Or A.效期 > Trunc(Sysdate))" & _
                " Group by A.药品ID Having Sum(Nvl(A.可用数量,0))<>0"
        ElseIf blnStock And Not (mlng西药房 = 0 And mlng成药房 = 0 And mlng中药房 = 0) Then
            '明确类别下未指定库房,不处理库存
            '不明确类别时都未指定库房,不处理库存
            strStock = _
                " Select A.药品ID,Sum(Nvl(A.可用数量,0)) as 库存" & _
                " From 药品库存 A,收费项目目录 B" & _
                " Where A.性质=1 And (Nvl(A.批次, 0)=0 Or A.效期 Is Null Or A.效期 > Trunc(Sysdate))" & _
                    " And A.库房ID=Decode(B.类别,'5',[12],'6',[13],'7',[14],Null)" & _
                    " And A.药品ID=B.ID And B.类别 IN('5','6','7')" & _
                " Group by A.药品ID Having Sum(Nvl(A.可用数量,0))<>0"
            'strStock = "" '优化
        End If
                
        If mstr输入 <> "" Then
            strInput = " And (A.编码 Like [6] And B.码类=[8] Or B.名称 Like [7] And B.码类=[8] Or B.简码 Like [7] And B.码类 IN([8],3))"
            If IsNumeric(mstr输入) Then                         '10,11.输入全是数字时只匹配编码'对于药品,则要匹配简码(码类为3的数字码)
                If Mid(gstrMatchMode, 1, 1) = "1" Then strInput = " And (A.编码 Like [6] And B.码类=[8] Or B.简码 Like [7] And B.码类=3)"
            ElseIf zlCommFun.IsCharAlpha(mstr输入) Then         '01,11.输入全是字母时只匹配简码
                If Mid(gstrMatchMode, 2, 1) = "1" Then strInput = " And B.简码 Like [7] And B.码类=[8]"
            ElseIf zlCommFun.IsCharChinese(mstr输入) Then
                strInput = " And B.名称 Like [7] And B.码类=[8]"
            End If
            
            strMain = _
                " Select Distinct A.ID,A.类别,A.编码,B.名称,B.简码," & _
                " A.计算单位,A.规格,A.产地,A.费用类型,A.说明,A.是否变价" & _
                " From 收费项目别名 B,收费项目目录 A" & _
                " Where A.ID=B.收费细目ID And A.服务对象 IN([9],3) And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL)" & _
                  IIF(str类别 <> "", " And A.类别=[4]", "") & IIF(mstr类别 <> "", " And Instr([5],A.类别)>0", "") & mstr特准项目 & strInput
        Else
            strMain = _
                " Select" & _
                " A.ID,A.类别,A.编码,Nvl(B.名称,A.名称) as 名称," & _
                " A.计算单位,A.规格,A.产地,A.费用类型,A.说明,A.是否变价" & _
                " From 收费项目别名 B,收费项目目录 A" & _
                " Where A.ID=B.收费细目ID(+) And B.码类(+)=1 And B.性质(+)=[10]" & _
                " And A.服务对象 IN([9],3) And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL)" & _
                 IIF(str类别 <> "", " And A.类别=[4]", "") & IIF(mstr类别 <> "", " And Instr([5],A.类别)>0", "") & mstr特准项目
        End If
            
        '药品单位
        If mint病人来源 = 1 Then
            str药房单位 = "门诊单位": str药房包装 = "门诊包装"
        Else
            str药房单位 = "住院单位": str药房包装 = "住院包装"
        End If
            
        If strStock = "" Then
            strSQL = _
                " Select A.ID,A.类别 as 类别ID,B.序号 as 顺序ID,B.名称 as 类别," & _
                    " A.编码,A.名称," & IIF(mstr输入 <> "", "A.简码,", "") & _
                    IIF(mbln药房单位, "D." & str药房单位, "A.计算单位") & " as 单位," & _
                    " A.规格,A.产地,A.费用类型,A.说明," & _
                    " Decode(A.是否变价,1,'时价',LTrim(To_Char(Sum(C.现价)" & _
                        IIF(mbln药房单位, "*Nvl(D." & str药房包装 & ",1)", "") & ",'9999990.00000'))) as 单价," & _
                    " NULL as 库存" & _
                " From (" & strMain & ") A,收费项目类别 B,收费价目 C,药品规格 D,诊疗项目目录 E" & _
                " Where A.类别=B.编码 And A.ID=C.收费细目ID" & _
                    " And A.ID=D.药品ID And D.药名ID=E.ID" & str分类ID & _
                    " And Sysdate Between C.执行日期 and Nvl(C.终止日期,To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " Group by A.ID,A.类别,B.序号,B.名称,A.编码,A.名称," & IIF(mstr输入 <> "", "A.简码,", "") & _
                    "A.规格,A.产地,A.费用类型,A.说明,A.是否变价," & _
                    IIF(mbln药房单位, "D." & str药房单位 & ",D." & str药房包装, "A.计算单位")
        Else
            '是否必须有库存:指定药房时根据系统参数是否要限制库存
            strTmp = " And A.ID=X.药品ID(+)"
            If InStr(mstrPrivs, "不检查库存") = 0 And gblnStock Then
                strTmp = strTmp & " And (" & _
                    " A.类别='5' And ([12]=0 Or X.药品ID Is Not Null)" & _
                    " Or A.类别='6' And ([13]=0 Or X.药品ID Is Not Null)" & _
                    " Or A.类别='7' And ([14]=0 Or X.药品ID Is Not Null)" & _
                    ")"
            End If
            strSQL = _
                " Select A.ID,A.类别 as 类别ID,B.序号 as 顺序ID,B.名称 as 类别," & _
                    " A.编码,A.名称," & IIF(mstr输入 <> "", "A.简码,", "") & _
                    IIF(mbln药房单位, "D." & str药房单位, "A.计算单位") & " as 单位," & _
                    " A.规格,A.产地,A.费用类型,A.说明," & _
                    " Decode(A.是否变价,1,'时价',LTrim(To_Char(Sum(C.现价)" & _
                        IIF(mbln药房单位, "*Nvl(D." & str药房包装 & ",1)", "") & ",'9999990.00000'))) as 单价," & _
                    " LTrim(To_Char(X.库存" & IIF(mbln药房单位, "/Nvl(D." & str药房包装 & ",1)", "") & ",'9999990.00000')) as 库存" & _
                " From (" & strMain & ") A,收费项目类别 B,收费价目 C," & _
                    " 药品规格 D,诊疗项目目录 E,(" & strStock & ") X" & _
                " Where A.类别=B.编码 And A.ID=C.收费细目ID" & _
                    " And A.ID=D.药品ID And D.药名ID=E.ID" & strTmp & str分类ID & _
                    " And Sysdate Between C.执行日期 and Nvl(C.终止日期,To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " Group by A.ID,A.类别,B.序号,B.名称,A.编码,A.名称," & IIF(mstr输入 <> "", "A.简码,", "") & _
                    "A.规格,A.产地,A.费用类型,A.说明,A.是否变价,X.库存," & _
                    IIF(mbln药房单位, "D." & str药房单位 & ",D." & str药房包装, "A.计算单位")
        End If
    End If
    
    '卫材收费项目部份
    '因为从诊疗中分类,所以单独处理。目前暂不管库存显示
    blnLoad = False
    If str类别 <> "" Then
        blnLoad = str类别 = "4"
    ElseIf mstr输入 <> "" Then
        If mstr类别 = "" Or InStr(mstr类别, "'4'") > 0 Then blnLoad = True
    End If
    If blnLoad Then
        If mstr输入 <> "" Then
            strInput = " And (A.编码 Like [6] Or B.名称 Like [7] Or B.简码 Like [7]) And B.码类=[8]"
            If IsNumeric(mstr输入) Then                         '10,11.输入全是数字时只匹配编码
                If Mid(gstrMatchMode, 1, 1) = "1" Then strInput = " And A.编码 Like [6] And B.码类=[8]"
            ElseIf zlCommFun.IsCharAlpha(mstr输入) Then         '01,11.输入全是字母时只匹配简码
                If Mid(gstrMatchMode, 2, 1) = "1" Then strInput = " And B.简码 Like [7] And B.码类=[8]"
            ElseIf zlCommFun.IsCharChinese(mstr输入) Then
                strInput = " And B.名称 Like [7] And B.码类=[8]"
            End If

            strMain = _
                " Select Distinct A.ID,A.类别,A.编码,B.名称,B.简码," & _
                " A.计算单位,A.规格,A.产地,A.费用类型,A.说明,A.是否变价" & _
                " From 收费项目别名 B,收费项目目录 A" & _
                " Where A.服务对象 IN([9],3) And A.ID=B.收费细目ID And A.类别='4'" & mstr特准项目 & strInput & _
                    " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL)"
        Else
            strMain = _
                " Select A.ID,A.类别,A.编码,A.名称,A.计算单位,A.规格,A.产地,A.费用类型,A.说明,A.是否变价" & _
                " From 收费项目目录 A" & _
                " Where A.服务对象 IN([9],3) And A.类别='4'" & mstr特准项目 & _
                    " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL)"
        End If
            
        strSQL = strSQL & IIF(strSQL = "", "", " Union ALL ") & _
            " Select A.ID,A.类别 as 类别ID,B.序号 as 顺序ID,B.名称 as 类别," & _
                " A.编码,A.名称," & IIF(mstr输入 <> "", "A.简码,", "") & _
                " A.计算单位 as 单位,A.规格,A.产地,A.费用类型,A.说明," & _
                " Decode(A.是否变价,1,'时价',LTrim(To_Char(Sum(C.现价),'9999990.00000'))) as 单价," & _
                " NULL as 库存" & _
            " From (" & strMain & ") A,收费项目类别 B,收费价目 C,材料特性 D,诊疗项目目录 E" & _
            " Where A.类别=B.编码 And A.ID=C.收费细目ID" & _
                " And A.ID=D.材料ID And D.诊疗ID=E.ID" & str分类ID & _
                " And Sysdate Between C.执行日期 and Nvl(C.终止日期,To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " Group by A.ID,A.类别,B.序号,B.名称,A.编码,A.名称," & IIF(mstr输入 <> "", "A.简码,", "") & _
                "A.规格,A.产地,A.费用类型,A.说明,A.是否变价,A.计算单位"
    End If
    
    '其它收费项目部份
    blnLoad = False
    If str类别 <> "" Then
        blnLoad = InStr(",4,5,6,7,", str类别) = 0
    Else
        strTmp = mstr类别
        strTmp = Replace(strTmp, "'4'", "")
        strTmp = Replace(strTmp, "'5'", "")
        strTmp = Replace(strTmp, "'6'", "")
        strTmp = Replace(strTmp, "'7'", "")
        strTmp = Trim(Replace(strTmp, ",", ""))
        If strTmp <> "" Or mstr类别 = "" Then blnLoad = True
    End If
    If blnLoad Then
        If mstr输入 <> "" Then
            strInput = " And (A.编码 Like [6] Or B.名称 Like [7] Or B.简码 Like [7]) And B.码类=[8]"
            If IsNumeric(mstr输入) Then                         '10,11.输入全是数字时只匹配编码
                If Mid(gstrMatchMode, 1, 1) = "1" Then strInput = " And A.编码 Like [6] And B.码类=[8]"
            ElseIf zlCommFun.IsCharAlpha(mstr输入) Then         '01,11.输入全是字母时只匹配简码
                If Mid(gstrMatchMode, 2, 1) = "1" Then strInput = " And B.简码 Like [7] And B.码类=[8]"
            ElseIf zlCommFun.IsCharChinese(mstr输入) Then
                strInput = " And B.名称 Like [7] And B.码类=[8]"
            End If
            
            strMain = _
                " Select Distinct A.ID,A.类别,A.编码,B.名称,B.简码," & _
                " A.计算单位,A.规格,A.产地,A.费用类型,A.说明,A.是否变价" & _
                " From 收费项目别名 B,收费项目目录 A" & _
                " Where A.ID=B.收费细目ID And A.服务对象 IN([9],3) And A.类别 Not IN('4','5','6','7','1') And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL)" & _
                 IIF(str类别 <> "", " And A.类别=[4]", "") & IIF(mstr类别 <> "", " And Instr([5],A.类别)>0", "") & str分类ID & mstr特准项目 & strInput

        Else
            strMain = _
                " Select A.ID,A.类别,A.编码,A.名称," & _
                " A.计算单位,A.规格,A.产地,A.费用类型,A.说明,A.是否变价" & _
                " From 收费项目目录 A" & _
                " Where A.服务对象 IN([9],3) And A.类别 Not IN('4','5','6','7','1') And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL)" & _
                 IIF(str类别 <> "", " And A.类别=[4]", "") & IIF(mstr类别 <> "", " And Instr([5],A.类别)>0", "") & str分类ID & mstr特准项目
        End If
        
        strSQL = strSQL & IIF(strSQL = "", "", " Union ALL ") & _
            " Select A.ID,A.类别 as 类别ID,B.序号 as 顺序ID,B.名称 as 类别," & _
            " A.编码,A.名称," & IIF(mstr输入 <> "", "A.简码,", "") & _
            " A.计算单位 as 单位,A.规格,A.产地,A.费用类型,A.说明," & _
            " Decode(A.是否变价,1,'变价',LTrim(To_Char(Sum(C.现价),'9999990.00000'))) as 单价,NULL as 库存" & _
            " From 收费价目 C,(" & strMain & ") A,收费项目类别 B" & _
            " Where A.类别=B.编码 And A.ID=C.收费细目ID" & _
            " And Sysdate Between C.执行日期+0 and Nvl(C.终止日期,To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " Group by A.ID,A.类别,B.序号,B.名称,A.编码,A.名称," & IIF(mstr输入 <> "", "A.简码,", "") & _
                "A.规格,A.产地,A.费用类型,A.说明,A.是否变价,A.计算单位"
    End If
    strSQL = strSQL & " Order by 顺序ID,编码" '使用类别序号排序
    
    On Error GoTo errH
    Screen.MousePointer = 11
    Set mrsItem = zlDatabase.OpenSQLRecord(strSQL, Me.Name, int类型, lng诊疗分类ID, lng收费分类ID, str类别, mstr类别, _
        UCase(mstr输入) & "%", mstrLike & UCase(mstr输入) & "%", mint简码 + 1, mint病人来源, IIF(gbln商品名, 3, 1), _
        lng药房ID, mlng西药房, mlng成药房, mlng中药房, 0, Replace(mstr类别, "'", ""))
    
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
        If InStr("单价,库存", vsItem.TextMatrix(0, i)) > 0 Then
            vsItem.ColAlignment(i) = 7
        Else
            vsItem.ColAlignment(i) = 1
        End If
        If vsItem.TextMatrix(0, i) Like "*ID" Then
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
    
    '建立分类卡片:有多类时且项目数较多时
    If blnClass And vsItem.Rows > 10 Then
        arrClass = Split(Mid(strClass, 2), ",")
        If UBound(arrClass) > 0 Then
            For i = 0 To UBound(arrClass)
                If i < 9 Then
                    '用Alt快捷键焦点无法处理
                    Set objTab = tabClass.Tabs.Add(, "_" & Left(arrClass(i), 1), Mid(arrClass(i), 2) & "(" & i + 1 & ")")
                Else
                    Set objTab = tabClass.Tabs.Add(, "_" & Left(arrClass(i), 1), Mid(arrClass(i), 2))
                End If
                objTab.Tag = Mid(arrClass(i), 2)
            Next
        End If
    End If
    
    '行号列宽度
    vsItem.ColWidth(0) = Me.TextWidth(vsItem.TextMatrix(vsItem.Rows - 1, 0) & " ")
    If vsItem.ColWidth(0) < 380 Then vsItem.ColWidth(0) = 380
    
    vsItem.Row = vsItem.FixedRows: vsItem.Col = vsItem.FixedCols
    Call vsItem_AfterRowColChange(-1, -1, vsItem.Row, vsItem.Col)
    vsItem.Redraw = flexRDDirect
        
    tabClass.Visible = tabClass.Tabs.Count > 1
    Call Form_Resize
    
    If mrsItem.RecordCount > 0 Then mrsItem.MoveFirst
    lblInfo.Caption = "当前选择：" & GetTreePath(tvw_s.SelectedItem) & tabClass.SelectedItem.Tag & "，共有 " & mrsItem.RecordCount & " 个项目"
        
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
