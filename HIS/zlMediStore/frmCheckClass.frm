VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCheckClass 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "药品分类选择"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6165
   Icon            =   "frmCheckClass.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin TabDlg.SSTab SSTab1 
      Height          =   5655
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   9975
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "药品分类(&1)"
      TabPicture(0)   =   "frmCheckClass.frx":6852
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "tvw分类"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "库房货位(&2)"
      TabPicture(1)   =   "frmCheckClass.frx":686E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "vsfStock"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "chk货位"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.CheckBox chk货位 
         Caption         =   "仅显示当前库房已分配的货位"
         Height          =   255
         Left            =   -74880
         TabIndex        =   4
         Top             =   480
         Value           =   1  'Checked
         Width           =   2655
      End
      Begin MSComctlLib.TreeView tvw分类 
         Height          =   5175
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   4395
         _ExtentX        =   7752
         _ExtentY        =   9128
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   476
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         ImageList       =   "img16"
         Appearance      =   1
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfStock 
         Height          =   4575
         Left            =   -74880
         TabIndex        =   5
         Top             =   840
         Width           =   4455
         _cx             =   7858
         _cy             =   8070
         Appearance      =   0
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
         BackColorFixed  =   8421504
         ForeColorFixed  =   16777215
         BackColorSel    =   14737632
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   3
         GridLines       =   0
         GridLinesFixed  =   0
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmCheckClass.frx":688A
         ScrollTrack     =   -1  'True
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
         ExplorerBar     =   5
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   1
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
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "退出(&E)"
      Height          =   350
      Left            =   4920
      TabIndex        =   1
      Top             =   960
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&Q)"
      Height          =   350
      Left            =   4920
      TabIndex        =   0
      Top             =   480
      Width           =   1100
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
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCheckClass.frx":68FA
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmCheckClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngDrugId As Long '库房货位
Private mstr货位 As String '货位
Private mstrSelRows As String
Private mstr药品id As String
Private mBlnClick As Boolean    '用来记录是否点击的确定按钮

Private Sub chk货位_Click()
    Load库房货位
End Sub

Private Sub cmdCancel_Click()
'    Call frmCheckCourseCard.get药品id("", 0)
    mBlnClick = False
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim str分类ID As String
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer
    Dim intItem As Integer
    
    On Error GoTo errHandl
    mBlnClick = True
    
    mstr货位 = ""
    mstr药品id = ""
    str分类ID = ""
    
    '取货位
    With vsfStock
        For intItem = 1 To .rows - 1
            If .Cell(flexcpChecked, intItem, .ColIndex("选择")) = flexChecked Then
                mstr货位 = .TextMatrix(intItem, .ColIndex("名称")) & "," & mstr货位
            End If
        Next
    End With
    
    '取得药品分类（不选分类表示所有分类）
    For intItem = 1 To tvw分类.Nodes.count
        If tvw分类.Nodes(intItem).Key = "Root" And tvw分类.Nodes(intItem).Checked = True Then
            
            gstrSQL = "select ID from 收费项目目录 where 类别 in('5','6','7') and  (撤档时间 IS NULL OR 撤档时间= to_date('3000-01-01','YYYY-MM-DD'))"
            Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "查询所有药品")
            Do While Not rsTemp.EOF
                mstr药品id = rsTemp!id & "," & mstr药品id
                rsTemp.MoveNext
            Loop
            Exit For
        ElseIf tvw分类.Nodes(intItem).Key <> "Root" And _
            tvw分类.Nodes(intItem).Key <> "_中成药" And _
            tvw分类.Nodes(intItem).Key <> "_中草药" And _
            tvw分类.Nodes(intItem).Key <> "_西成药" And _
            tvw分类.Nodes(intItem).Checked Then
            str分类ID = str分类ID & "," & Mid(tvw分类.Nodes(intItem).Key, 2)
        End If
    Next

    If str分类ID <> "" Then
        str分类ID = Mid(str分类ID, 2)
    
        gstrSQL = "Select 药品id" & _
                  "  From 药品规格 " & _
                   " Where 药名id In (Select ID" & _
                    "               From 诊疗项目目录 A " & _
                    "               Where a.分类id  in (select * from Table(Cast(f_Num2list([1]) As Zltools.t_Numlist))))"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, str分类ID)
        If rsTemp Is Nothing Then
            Unload Me
            Exit Sub
        Else
            mstr药品id = ""
            For i = 1 To rsTemp.RecordCount
                mstr药品id = rsTemp!药品ID & "," & mstr药品id
                rsTemp.MoveNext
            Next
        End If
    End If
    
    Unload Me
    Exit Sub
errHandl:
If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    '药品分类
    Dim objnode As Node
    Dim rsTemp As ADODB.Recordset
    
    gstrSQL = "Select Level as 层,ID,上级ID,名称,DECODE(类型,1,'西成药',2,'中成药','中草药') As 材质 " & _
        " From 诊疗分类目录" & _
        " Where 类型 in (1,2,3)" & _
        " Start With 上级ID is NULL Connect by Prior ID=上级ID" & _
        " Order by Level,编码"
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, "提取药品分类")

    Set objnode = tvw分类.Nodes.Add(, , "Root", "所有用途", 1)
    Set objnode = tvw分类.Nodes.Add("Root", 4, "_西成药", "西成药", 1)
    Set objnode = tvw分类.Nodes.Add("Root", 4, "_中草药", "中草药", 1)
    Set objnode = tvw分类.Nodes.Add("Root", 4, "_中成药", "中成药", 1)

    Do While Not rsTemp.EOF
        If rsTemp!层 = 1 Then
            Set objnode = tvw分类.Nodes.Add("_" & rsTemp!材质, 4, "_" & rsTemp!id, rsTemp!名称, 1)
        Else
            Set objnode = tvw分类.Nodes.Add("_" & rsTemp!上级ID, 4, "_" & rsTemp!id, rsTemp!名称, 1)
        End If
        rsTemp.MoveNext
    Loop
    tvw分类.Nodes("Root").Selected = True
    tvw分类.Nodes("Root").Expanded = True
    
    Call Load库房货位
End Sub

Private Sub Load库房货位()
    Dim rsData As ADODB.Recordset
    Dim lngRow As Long
    Dim intLevel As Integer
    
    On Error GoTo errHandle
    If chk货位.Value = 1 Then
        gstrSQL = "Select a.编码, a.名称 From 药品库房货位 A " & _
            " Where a.库房id = [1] And Exists (Select 1 From 药品货位对照 B Where b.库房id = a.库房id And b.货位id = a.Id) " & _
            " Order By 名称 "
    Else
        gstrSQL = "Select 编码, 名称 From  药品库房货位 Where 库房id = [1] Order By 名称 "
    End If
    
    Set rsData = zldatabase.OpenSQLRecord(gstrSQL, "提取所有药品库房货位", mlngDrugId)
    
    If rsData.RecordCount = 0 Then
        vsfStock.rows = 1
        Exit Sub
    End If
    
    With vsfStock
        .Redraw = flexRDNone
        .rows = 1
        
        Do While Not rsData.EOF
            .rows = .rows + 1
            
            .TextMatrix(.rows - 1, .ColIndex("选择")) = 0
            .TextMatrix(.rows - 1, .ColIndex("编码")) = rsData!编码
            .TextMatrix(.rows - 1, .ColIndex("名称")) = rsData!名称
            
            rsData.MoveNext
        Loop
        
        .Redraw = flexRDDirect
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetSubItem(ByVal str上级编码 As String, ByVal rsData As ADODB.Recordset)
    '用递归算法找树表的子项目
    Dim rsClone As ADODB.Recordset
    
    Set rsClone = rsData.Clone
    
    rsClone.Filter = "上级='" & str上级编码 & "'"
    rsClone.Sort = "名称"
    
    '没找到下一级时一定要退出
    If rsClone.RecordCount = 0 Then Exit Sub
    
    With vsfStock
        .Redraw = flexRDNone
        
        Do While Not rsClone.EOF
            .rows = .rows + 1
            
            .TextMatrix(.rows - 1, .ColIndex("级数")) = rsClone!级数 + 1
            .TextMatrix(.rows - 1, .ColIndex("上级")) = rsClone!上级
            .TextMatrix(.rows - 1, .ColIndex("名称")) = rsClone!名称
            .TextMatrix(.rows - 1, .ColIndex("选择")) = 0
            .TextMatrix(.rows - 1, .ColIndex("编码")) = rsClone!编码
            
            '找下一级的项目
            Call GetSubItem(rsClone!编码, rsData)
            
            rsClone.MoveNext
        Loop
        
        .Redraw = flexRDDirect
    End With
End Sub

Public Sub ShowME(ByVal frmCard As frmNewCheckCourseCard, ByVal lng库房ID As Long, ByRef str货位 As String, ByRef str药品ID As String, ByRef int退出 As Integer)
    mBlnClick = False
    
    mlngDrugId = lng库房ID
    Me.Show vbModal, frmCard
    
    str货位 = mstr货位
    str药品ID = mstr药品id
    
    If mBlnClick = True Then
        int退出 = 1
    Else
        int退出 = 0
    End If
End Sub

Private Sub tvw分类_NodeCheck(ByVal Node As MSComctlLib.Node)
    CheckNode Node, Node.Checked
End Sub

Private Function CheckNode(ByVal Node As Object, blnCheck As Boolean)
    Dim intIdx As Integer

    If Node.Children > 0 Then
        Set Node = Node.Child
        Do While Not Node Is Nothing
            Node.Checked = blnCheck
            If Node.Children > 0 Then
                CheckNode Node, blnCheck
            End If
            Set Node = Node.Next
        Loop
    Else
        Node.Checked = blnCheck
    End If
End Function





Private Sub CheckItem(ByVal str编码 As String, ByVal intChecked As Integer)
    Dim lngRows As Long
    
    With vsfStock
        If str编码 = "所有" Then
            '当前是根节点，所有库房设置为全选或全不选
            For lngRows = 2 To .rows - 1
                .Cell(flexcpChecked, lngRows, .ColIndex("选择")) = intChecked
            Next
        Else
            '当前是子节点，需要递归处理下级子节点
            For lngRows = 2 To .rows - 1
                If .TextMatrix(lngRows, .ColIndex("上级")) = str编码 Then
                    .Cell(flexcpChecked, lngRows, .ColIndex("选择")) = intChecked
                    
                    Call CheckItem(.TextMatrix(lngRows, .ColIndex("编码")), intChecked)
                End If
            Next
        End If
    End With
End Sub


Private Sub vsfStock_BeforeMouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single, Cancel As Boolean)
    Dim lngRow As Long
    Dim strFlag As String

    With vsfStock
        If .MouseRow <= 0 Then Exit Sub
        If .MouseCol <> .ColIndex("选择") Then Exit Sub
        
        mstrSelRows = ""

        For lngRow = 1 To .rows - 1
            If .IsSelected(lngRow) Then
                mstrSelRows = IIf(mstrSelRows = "", "", mstrSelRows & ",") & lngRow
            End If
        Next
    End With
End Sub


Private Sub vsfStock_Click()
    Dim IntCheck As Integer
    Dim lngRow As Long
    Dim lngRows As Long
    
    With vsfStock
        If .MouseRow <= 0 Then Exit Sub
        If .MouseCol <> .ColIndex("选择") Then Exit Sub
        
        lngRow = .MouseRow
        
        IntCheck = .Cell(flexcpChecked, lngRow, .ColIndex("选择"))
        
'        Call CheckItem(.TextMatrix(lngRow, .ColIndex("编码")), IntCheck)

        If InStr(1, mstrSelRows, ",") > 0 Then
            '多选处理
            For lngRows = 1 To .rows - 1
                If InStr(1, "," & mstrSelRows & ",", "," & lngRows & ",") > 0 Then
                    .Cell(flexcpChecked, lngRows, .ColIndex("选择")) = IntCheck
                       
                End If
            Next
            
            mstrSelRows = ""
        Else
            '单选时处理当前节点
            .Cell(flexcpChecked, lngRow, .ColIndex("选择")) = IntCheck
        End If
        
    End With
End Sub
