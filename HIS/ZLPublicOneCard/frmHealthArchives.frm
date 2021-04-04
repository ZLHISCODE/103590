VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmHealthArchives 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "居民健康档案信息"
   ClientHeight    =   10485
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12000
   Icon            =   "frmHealthArchives.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10485
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdExit 
      Caption         =   "退出(&X)"
      Height          =   350
      Left            =   10485
      TabIndex        =   4
      Top             =   360
      Width           =   1100
   End
   Begin VB.PictureBox picPatiInfor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   9495
      Left            =   210
      ScaleHeight     =   9465
      ScaleWidth      =   9675
      TabIndex        =   0
      Top             =   210
      Width           =   9705
      Begin VSFlex8Ctl.VSFlexGrid vsGrid 
         Height          =   8850
         Left            =   -15
         TabIndex        =   1
         Top             =   -15
         Width           =   9765
         _cx             =   17224
         _cy             =   15610
         Appearance      =   0
         BorderStyle     =   0
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
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   16777215
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483639
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   23
         Cols            =   9
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   350
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmHealthArchives.frx":0442
         ScrollTrack     =   0   'False
         ScrollBars      =   0
         ScrollTips      =   0   'False
         MergeCells      =   1
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
         Begin VB.PictureBox picPhoto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   1740
            Left            =   6645
            ScaleHeight     =   1710
            ScaleWidth      =   3030
            TabIndex        =   2
            Top             =   -15
            Width           =   3060
            Begin VB.Image imgPhoto 
               Height          =   435
               Left            =   1185
               Stretch         =   -1  'True
               Top             =   765
               Width           =   315
            End
         End
      End
   End
   Begin XtremeSuiteControls.TaskPanel wndTaskPanel 
      Height          =   10290
      Left            =   120
      TabIndex        =   3
      Top             =   60
      Width           =   10125
      _Version        =   589884
      _ExtentX        =   17859
      _ExtentY        =   18150
      _StockProps     =   64
      VisualTheme     =   6
      ItemLayout      =   2
      HotTrackStyle   =   1
   End
End
Attribute VB_Name = "frmHealthArchives"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const M_IDX_TP_BASE = 100
Private mlng病人ID As Long
Private mrsInfor As ADODB.Recordset '病人基本信息
Private mrsOtherCertificate As ADODB.Recordset
Private mrsDrug As ADODB.Recordset
Private mrsBacterin As ADODB.Recordset
Private mblnUnLoad As Boolean
Private mcnOracle As ADODB.Connection
Private mobjDataBase As clsDataBase
Public Sub zlShowHealthArchives(ByVal frmMain As Object, ByVal lng病人ID As Long, _
    Optional cnOracle As ADODB.Connection)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示档案信息
    '入参:lng病人ID-病人ID
    '     cnOracle-数据库连接
    '编制:刘兴洪
    '日期:2012-12-14 13:34:29
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mlng病人ID = lng病人ID
    Set mcnOracle = cnOracle
    If zlGetOneDataBase(cnOracle, mobjDataBase) = False Then Exit Sub
    
    If frmMain Is Nothing Then
        Me.Show 1
    Else
        Me.Show 1, frmMain
    End If
End Sub
Private Function LoadPatiInfor() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载病人信息
    '返回:成在病人信息,返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-12-14 13:38:09
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String
    Err = 0: On Error GoTo errHandle
    
    
    '82072:李南春,2015/1/23,血型和RH取就诊ID 为null的记录
    strSql = "" & _
    "  Select a.病人id,a.姓名,to_char(a.出生日期,'yyyy-mm-dd') as 出生日期,a.性别,a.年龄,a.民族,a.婚姻状况,a.学历,a.家庭电话,a.职业,max(a.身份证号) as 身份证号, " & _
    "          max(a.联系人姓名) as 联系人姓名1,max(a.联系人关系) as 联系人关系1,max(a.联系人电话) as 联系人电话1, " & _
    "          max(a.户口地址) as 户口地址,max(a.家庭地址) as 家庭地址, " & _
    "          max(decode(b.信息名,'联系人姓名1',b.信息值,'')) as 联系人姓名2,max(decode(b.信息名,'联系人关系1',b.信息值,'')) as  联系人关系2,max(decode(b.信息名,'联系人电话1',b.信息值,''))   as 联系人电话2, " & _
    "          max(decode(b.信息名,'联系人姓名2',b.信息值,'')) as 联系人姓名3,max(decode(b.信息名,'联系人关系2',b.信息值,'')) as  联系人关系3,max(decode(b.信息名,'联系人电话2',b.信息值,''))   as 联系人电话3, " & _
    "          max(decode(b.信息名,'新农合(卡)号',b.信息值,'')) as  新农合号, max(decode(b.信息名,'健康档案编号',b.信息值,'')) as  健康档案编号, " & _
    "          max(decode(b.信息名,'医疗费用支付方式',b.信息值,'')) as  费用支付方式, " & _
    "          max(decode(b.就诊id,null,decode(b.信息名,'ABO',b.信息值,'血型',b.信息值,''),'')) as  ABO血型, " & _
    "          max(decode(b.就诊id,null,decode(b.信息名,'RH',b.信息值,''),'')) as  RH," & _
    "          max(decode(b.信息名,'医学警示',b.信息值,'')) as  医学警示," & _
    "          max(decode(b.信息名,'其他医学警示',b.信息值,'')) as  其他医学警示" & _
    "   From 病人信息 A,病人信息从表 B  " & _
    "   Where  a.病人ID=b.病人ID(+) and a.病人ID=[1] " & _
    "   Group by a.病人ID,a.姓名,a.出生日期,a.性别,a.年龄,a.民族,a.婚姻状况,a.学历,a.家庭电话,a.职业"
    
    Set mrsInfor = mobjDataBase.OpenSQLRecord(strSql, Me.Caption, mlng病人ID)
    If mrsInfor.RecordCount = 0 Then
        MsgBox "当前病人信息不存在，请检查！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '读取其他证件
    strSql = "Select a.信息名 as 证件类型, a.信息值 as 证件号码 From 病人信息从表 A, 证件类型 B Where a.病人id = [1] And a.信息名 = b.名称  Order by 信息名"
    Set mrsOtherCertificate = mobjDataBase.OpenSQLRecord(strSql, Me.Caption, mlng病人ID)
    
    '读取过敏药物
    strSql = "  Select 过敏药物,过敏反应 from 病人过敏药物 where 病人ID=[1] order by 过敏药物"
    Set mrsDrug = mobjDataBase.OpenSQLRecord(strSql, Me.Caption, mlng病人ID)
    
    '读取疫苗接种情况
    strSql = " " & _
    "   Select 病人id, To_char(接种时间, 'yyyy-mm-dd hh24:mi') As 接种时间, 接种名称 " & _
    "   From 病人免疫记录 " & _
    "   Where 病人id = [1] " & _
    "   Order By 接种时间"
    Set mrsBacterin = mobjDataBase.OpenSQLRecord(strSql, Me.Caption, mlng病人ID)
    LoadPatiInfor = True
    Exit Function
errHandle:
    If mobjDataBase.ErrCenter() = 1 Then
        Resume
    End If
    
End Function

Private Sub InitTaskPancel()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化InitTaskPancel
    '编制:刘兴洪
    '日期:2012-12-13 17:48:22
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim tkpGroup As TaskPanelGroup
    Dim Item As TaskPanelGroupItem
    'Call wndTaskPanel.SetGroupInnerMargins(2, 0, 2, 0)
    
    Call wndTaskPanel.SetGroupOuterMargins(2, -10, 2, -10)
      Call wndTaskPanel.SetMargins(2, 16, 2, 10, 30)
    wndTaskPanel.HotTrackStyle = xtpTaskPanelHighlightItem
    Set tkpGroup = wndTaskPanel.Groups.Add(M_IDX_TP_BASE, "病人健康档案信息")
    Set Item = tkpGroup.Items.Add(M_IDX_TP_BASE, "", xtpTaskItemTypeControl)
    Set Item.Control = picPatiInfor
    picPatiInfor.BackColor = Item.BackColor
    tkpGroup.Expandable = False
    wndTaskPanel.Reposition
    wndTaskPanel.DrawFocusRect = True
End Sub

Private Sub InitGrid()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化网格
    '编制:刘兴洪
    '日期:2012-12-13 18:35:13
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, r As Long
    
    On Error GoTo errHandle
    
    With vsGrid
        .Clear
        .MergeCells = flexMergeFree
        .Rows = 27: .Cols = 9
        .RowHeightMin = 350
        .FixedRows = 0: .FixedCols = 0
        For i = 0 To .Cols - 1
            .ColAlignment(i) = flexAlignCenterCenter
            .FixedAlignment(i) = flexAlignCenterCenter
        Next
        .RowHidden(0) = True
        .TextMatrix(1, 1) = "姓名"
        .TextMatrix(1, 2) = "姓名"
        For i = 3 To 6
            .TextMatrix(1, i) = NVL(mrsInfor!姓名, " ")
        Next
        .Cell(flexcpAlignment, 1, 3, 1, 6) = 1
        .Cell(flexcpForeColor, 1, 3, 1, 6) = vbBlue
        .TextMatrix(2, 1) = "出生日期"
        .TextMatrix(2, 2) = "出生日期"
        For i = 3 To 6
            .TextMatrix(2, i) = NVL(mrsInfor!出生日期, "  ")
        Next
        .Cell(flexcpAlignment, 2, 3, 2, 6) = 1
        .Cell(flexcpForeColor, 2, 3, 2, 6) = vbBlue
        
        .TextMatrix(3, 1) = "性别"
        .TextMatrix(3, 2) = "性别"
        .TextMatrix(3, 3) = NVL(mrsInfor!性别, "   ")
        .Cell(flexcpAlignment, 3, 3, 3, 3) = 1
        .Cell(flexcpForeColor, 3, 3, 3, 3) = vbBlue
        
        .TextMatrix(3, 4) = "民族"
        .TextMatrix(3, 5) = NVL(mrsInfor!民族, "   ")
        .Cell(flexcpForeColor, 3, 5, 3, 5) = vbBlue
        .TextMatrix(3, 6) = .TextMatrix(3, 5)
        .Cell(flexcpForeColor, 3, 5, 3, 6) = vbBlue
        .Cell(flexcpAlignment, 3, 5, 3, 6) = 1
        
                
        .TextMatrix(4, 1) = "婚姻状况"
        .TextMatrix(4, 2) = "婚姻状况"
        .TextMatrix(4, 3) = NVL(mrsInfor!婚姻状况)
        .Cell(flexcpForeColor, 4, 3, 4, 3) = vbBlue
        .Cell(flexcpAlignment, 4, 3, 4, 3) = 1
        .TextMatrix(4, 4) = "文化程度"
        .TextMatrix(4, 5) = NVL(mrsInfor!学历, "  ")
        .TextMatrix(4, 6) = .TextMatrix(4, 5)
        .Cell(flexcpAlignment, 4, 5, 4, 6) = 1
        .Cell(flexcpForeColor, 4, 5, 4, 6) = vbBlue
        .Cell(flexcpAlignment, 5, 3, 8, 3) = 1

        
        .TextMatrix(5, 1) = "本人电话"
        .TextMatrix(5, 2) = "本人电话"
        .TextMatrix(5, 3) = NVL(mrsInfor!家庭电话, " ")
        .Cell(flexcpForeColor, 5, 3, 5, 3) = vbBlue
        .TextMatrix(5, 4) = "职业"
        .TextMatrix(5, 5) = NVL(mrsInfor!职业, "   ")
        .TextMatrix(5, 6) = .TextMatrix(5, 5)
        .Cell(flexcpAlignment, 5, 5, 5, 6) = 1
        .Cell(flexcpForeColor, 5, 5, 5, 6) = vbBlue
                
        .TextMatrix(6, 1) = "联系人"
        .TextMatrix(7, 1) = "联系人"
        .TextMatrix(8, 1) = "联系人"
        
        .TextMatrix(6, 2) = "姓名1"
        .TextMatrix(6, 3) = NVL(mrsInfor!联系人姓名1, "")
        .Cell(flexcpAlignment, 6, 3, 6, 3) = 1
        .Cell(flexcpForeColor, 6, 3, 6, 3) = vbBlue
        
        .TextMatrix(7, 2) = "姓名2"
        .TextMatrix(7, 3) = NVL(mrsInfor!联系人姓名2, "  ")
        .Cell(flexcpAlignment, 7, 3, 7, 3) = 1
        .Cell(flexcpForeColor, 7, 3, 7, 3) = vbBlue
        
        .TextMatrix(8, 2) = "姓名3"
        .TextMatrix(8, 3) = NVL(mrsInfor!联系人姓名3, "")
        .Cell(flexcpForeColor, 8, 3, 8, 3) = vbBlue
        .Cell(flexcpAlignment, 8, 3, 8, 3) = 1
        .Cell(flexcpAlignment, 6, 5, 8, 5) = 1
        .Cell(flexcpForeColor, 6, 5, 8, 5) = vbBlue
        
        
        .TextMatrix(6, 4) = "关系1"
        .TextMatrix(6, 5) = NVL(mrsInfor!联系人关系1, "")
        .TextMatrix(7, 4) = "关系2"
        .TextMatrix(7, 5) = NVL(mrsInfor!联系人关系2, " ")
        .TextMatrix(8, 4) = "关系3"
        .TextMatrix(8, 5) = NVL(mrsInfor!联系人关系3, "")
        
        .TextMatrix(6, 6) = "电话1"
        .TextMatrix(7, 6) = "电话2"
        .TextMatrix(8, 6) = "电话3"
        
        .Cell(flexcpAlignment, 6, 7, 8, 8) = 1
        .Cell(flexcpForeColor, 6, 7, 8, 8) = vbBlue
        For i = 7 To 8
            .TextMatrix(6, i) = NVL(mrsInfor!联系人电话1, " ")
        Next
        For i = 7 To 8
            .TextMatrix(7, i) = NVL(mrsInfor!联系人电话2, "  ")
        Next
        For i = 7 To 8
            .TextMatrix(8, i) = NVL(mrsInfor!联系人电话3, "   ")
        Next
        
        For i = 9 To 12
            .TextMatrix(i, 1) = "身份标识"
            .TextMatrix(i, 2) = "身份标识"
        Next
                        
        .TextMatrix(9, 3) = "身份证"
        .Cell(flexcpAlignment, 9, 4, 9, .Cols - 1) = 1
        .Cell(flexcpForeColor, 9, 4, 9, .Cols - 1) = vbBlue
        For i = 4 To .Cols - 1
            .TextMatrix(9, i) = NVL(mrsInfor!身份证号, " ")
        Next
        
        .TextMatrix(10, 3) = "其他证件"
        .TextMatrix(10, 6) = "证件号码"
        .Cell(flexcpAlignment, 10, 4, 10, 5) = 1
        .Cell(flexcpAlignment, 10, 7, 10, 8) = 1
        .Cell(flexcpForeColor, 10, 4, 10, 5) = vbBlue
        .Cell(flexcpForeColor, 10, 7, 10, 8) = vbBlue
        If mrsOtherCertificate.RecordCount > 0 Then
            .TextMatrix(10, 4) = NVL(mrsOtherCertificate!证件类型)
            .TextMatrix(10, 5) = NVL(mrsOtherCertificate!证件类型)
            
            .TextMatrix(10, 7) = NVL(mrsOtherCertificate!证件号码)
            .TextMatrix(10, 8) = .TextMatrix(10, 7)
        Else
            .TextMatrix(10, 4) = "     "
            .TextMatrix(10, 5) = "     "
            For i = 7 To .Cols - 1
                .TextMatrix(10, i) = "  "
            Next
        End If
                
        .TextMatrix(11, 3) = "新农合证(卡)号"
        .Cell(flexcpAlignment, 11, 4, 11, 5) = 1
        .Cell(flexcpForeColor, 11, 4, 11, 5) = vbBlue
        For i = 4 To 5
            .TextMatrix(11, i) = NVL(mrsInfor!新农合号, "  ")
        Next
        
        .TextMatrix(11, 6) = "健康档案编号"
        .Cell(flexcpAlignment, 11, 7, 11, 8) = 1
        .Cell(flexcpForeColor, 11, 7, 11, 8) = vbBlue
        For i = 7 To .Cols - 1
            .TextMatrix(11, i) = NVL(mrsInfor!健康档案编号, "   ")
        Next
        
        .TextMatrix(12, 1) = "户籍地址"
        .TextMatrix(12, 2) = "户籍地址"
        .TextMatrix(12, 3) = NVL(mrsInfor!户口地址, " ")
        .Cell(flexcpForeColor, 12, 3, 12, .Cols - 1) = vbBlue
        .Cell(flexcpAlignment, 12, 3, 12, .Cols - 1) = 1
        For i = 4 To .Cols - 1
            .TextMatrix(12, i) = .TextMatrix(12, 3)
        Next
        
        .TextMatrix(13, 1) = "居住地址"
        .TextMatrix(13, 2) = "居住地址"
        .TextMatrix(13, 3) = NVL(mrsInfor!家庭地址, "  ")
        .Cell(flexcpAlignment, 13, 3, 13, .Cols - 1) = 1
        .Cell(flexcpForeColor, 13, 3, 13, .Cols - 1) = vbBlue
        For i = 4 To .Cols - 1
            .TextMatrix(13, i) = .TextMatrix(13, 3)
        Next
        
        .TextMatrix(14, 1) = "医疗费用支付方式"
        .TextMatrix(14, 2) = "医疗费用支付方式"
        .TextMatrix(14, 3) = NVL(mrsInfor!费用支付方式, " ")
        .RowHeight(14) = 600
        .Cell(flexcpAlignment, 14, 3, 14, .Cols - 1) = 1
        .Cell(flexcpForeColor, 14, 3, 14, .Cols - 1) = vbBlue
        For i = 4 To .Cols - 1
            .TextMatrix(14, i) = .TextMatrix(14, 3)
        Next
                
        For i = 1 To 14
            .TextMatrix(i, 0) = "身份识别数据"
        Next
        For i = 15 To .Rows - 1
            .TextMatrix(i, 0) = "基础健康数据"
        Next
        
        .TextMatrix(15, 1) = "生物标识"
        .TextMatrix(15, 2) = "生物标识"
        .TextMatrix(15, 3) = "ABO血型"
        .TextMatrix(15, 4) = NVL(mrsInfor!ABO血型, "  ")
        .TextMatrix(15, 5) = .TextMatrix(15, 4)
        .Cell(flexcpAlignment, 15, 4, 15, 5) = 1
        .Cell(flexcpForeColor, 15, 4, 15, 5) = vbBlue
        
        .TextMatrix(15, 6) = "RH"
        .TextMatrix(15, 7) = NVL(mrsInfor!RH, "    ")
        .TextMatrix(15, 8) = .TextMatrix(15, 7)
        .Cell(flexcpAlignment, 15, 7, 15, 8) = 1
        .Cell(flexcpForeColor, 15, 7, 15, 8) = vbBlue
        
       For r = 16 To 21
            .TextMatrix(r, 1) = "医院警示"
            .TextMatrix(r, 2) = "医院警示"
        Next
        .TextMatrix(16, 3) = NVL(mrsInfor!医学警示, " ")
        .TextMatrix(16, 3) = IIf(Trim(.TextMatrix(16, 3)) = "", "", Trim(.TextMatrix(16, 3)) & ";") & NVL(mrsInfor!其他医学警示, " ")
        .Cell(flexcpAlignment, 16, 3, 16, .Cols - 1) = 1
        .Cell(flexcpForeColor, 16, 3, 16, .Cols - 1) = vbBlue
        
        For i = 4 To .Cols - 1
            .TextMatrix(16, i) = .TextMatrix(16, 3)
        Next
        .RowHeight(16) = 600
        
        r = 17
        .Cell(flexcpBackColor, r, 3, r, .Cols - 1) = &HFFC0C0
        .Cell(flexcpBackColor, r + 1, 3, r + 1, .Cols - 1) = &H8000000F
        For i = 3 To .Cols - 1
            .TextMatrix(r, i) = "过敏药物"
        Next
        .TextMatrix(18, 3) = "药物名称"
        .TextMatrix(18, 4) = "药物名称"
        .TextMatrix(18, 5) = "药物反应"
        .TextMatrix(18, 6) = "药物反应"
        .TextMatrix(18, 7) = "药物反应"
        .TextMatrix(18, 8) = "药物反应"
        .Cell(flexcpAlignment, 19, 3, 21, 8) = 1
        .Cell(flexcpForeColor, 19, 3, 21, .Cols - 1) = vbBlue
        
        r = 19
        Do While Not mrsDrug.EOF
            If r > 21 Then Exit Do
            .TextMatrix(r, 3) = NVL(mrsDrug!过敏药物) & Space(r - 19 + 1)
            .TextMatrix(r, 4) = .TextMatrix(r, 3)
            .TextMatrix(r, 5) = NVL(mrsDrug!过敏反应) & Space(r - 19 + 1)
            .TextMatrix(r, 6) = .TextMatrix(r, 5)
            .TextMatrix(r, 7) = .TextMatrix(r, 5)
            .TextMatrix(r, 8) = .TextMatrix(r, 5)
            r = r + 1
            mrsDrug.MoveNext
        Loop
        If r <= 21 Then
            For i = r To 21
                .TextMatrix(i, 3) = Space(i - 19 + 1)
                .TextMatrix(i, 4) = Space(i - 19 + 1)
                
                .TextMatrix(i, 5) = Space(i - 19 + 2)
                .TextMatrix(i, 6) = Space(i - 19 + 2)
                .TextMatrix(i, 7) = Space(i - 19 + 2)
                .TextMatrix(i, 8) = Space(i - 19 + 2)
            Next
        End If
        For r = 22 To 26
            .TextMatrix(r, 1) = "免疫接种"
            .TextMatrix(r, 2) = "免疫接种"
        Next
        .Cell(flexcpAlignment, 23, 3, .Rows - 1, 4) = 1
        .Cell(flexcpAlignment, 23, 6, .Rows - 1, 7) = 1
        .Cell(flexcpForeColor, 23, 3, .Rows - 1, .Cols - 1) = vbBlue
        r = 22
        .Cell(flexcpBackColor, r, 3, r, .Cols - 1) = &H8000000F
        .TextMatrix(22, 3) = "接种名称"
        .TextMatrix(22, 4) = "接种名称"
        .TextMatrix(22, 5) = "接种日期"
        
        .TextMatrix(22, 6) = "接种名称"
        .TextMatrix(22, 7) = "接种名称"
        .TextMatrix(22, 8) = "接种日期"
        r = 23
        i = 0
        Do While Not mrsBacterin.EOF
            If r > .Rows - 1 Then Exit Do
            If i = 0 Then
                .TextMatrix(r, 3) = NVL(mrsBacterin!接种名称) & Space(r - 19 + 1)
                .TextMatrix(r, 4) = .TextMatrix(r, 3)
                .TextMatrix(r, 5) = NVL(mrsBacterin!接种时间) & Space(r - 19 + 1)
            Else
                .TextMatrix(r, 6) = NVL(mrsBacterin!接种名称) & Space(r - 19 + 1)
                .TextMatrix(r, 7) = .TextMatrix(r, 6)
                .TextMatrix(r, 8) = NVL(mrsBacterin!接种时间) & Space(r - 19 + 1)
            End If
            If i Mod 2 <> 0 Then
                r = r + 1
                i = 0
            Else
                i = 1
            End If
            mrsBacterin.MoveNext
        Loop
        
        For i = r To .Rows - 1
            If Trim(.TextMatrix(i, 3)) = "" Then
                .TextMatrix(i, 3) = Space(i - 19 + 1)
                .TextMatrix(i, 4) = Space(i - 19 + 1)
            End If
            If Trim(.TextMatrix(i, 6)) = "" Then
                .TextMatrix(i, 6) = Space(i - 19 + 1)
                .TextMatrix(i, 7) = Space(i - 19 + 1)
            End If
        Next
        For i = 0 To .Rows - 1
            .MergeRow(i) = True
        Next
        For i = 0 To .Cols - 1
            .MergeCol(i) = True
        Next
        .WordWrap = True
    End With
    

    Exit Sub
errHandle:
    If mobjDataBase.ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If mblnUnLoad Then Unload Me: Exit Sub
End Sub

Private Sub Form_Load()
    mblnUnLoad = Not LoadPatiInfor
    If mblnUnLoad Then Exit Sub
    
    Call InitGrid
    Call LoadPhoto
    Call InitTaskPancel
    Call picPatiInfor_Resize
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Err = 0: On Error Resume Next
    If Not mrsInfor Is Nothing Then Set mrsInfor = Nothing
    If Not mrsOtherCertificate Is Nothing Then Set mrsOtherCertificate = Nothing
    If Not mrsDrug Is Nothing Then Set mrsDrug = Nothing
    If Not mrsBacterin Is Nothing Then Set mrsBacterin = Nothing
    If Not mobjDataBase Is Nothing Then Set mobjDataBase = Nothing
    If Not mcnOracle Is Nothing Then Set mcnOracle = Nothing
End Sub

Private Sub picPatiInfor_Resize()
    Err = 0: On Error Resume Next
    With picPatiInfor
        vsGrid.Left = .ScaleLeft
        vsGrid.Top = .ScaleTop
        vsGrid.Width = .ScaleWidth + 15
        vsGrid.Height = .ScaleHeight
    End With
End Sub
Private Sub LoadPhoto()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载像片
    '编制:刘兴洪
    '日期:2012-12-14 16:01:43
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strTempFile As String
    Dim objTemp As clsDataBase
    Dim objDatabase As Object
    
     
    '显示照片
    picPhoto.Cls
    strTempFile = mobjDataBase.ReadLob(glngSys, 27, mlng病人ID)
    imgPhoto.Picture = LoadPicture(strTempFile)
    '删除该临时文件
    Kill strTempFile
    imgPhoto.Left = picPhoto.ScaleLeft
    imgPhoto.Top = picPhoto.ScaleTop
    imgPhoto.Width = picPhoto.ScaleWidth
    imgPhoto.Height = picPhoto.ScaleHeight
    Set objDatabase = Nothing
End Sub
