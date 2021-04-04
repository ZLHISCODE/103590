VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "*\Azl9PacsControl\zl9PacsControl.vbp"
Begin VB.UserControl ucReportSegment 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6585
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LockControls    =   -1  'True
   ScaleHeight     =   9000
   ScaleWidth      =   6585
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1440
      Top             =   0
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
            Picture         =   "ucReportSegment.ctx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ucReportSegment.ctx":06FA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picBack 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8415
      Left            =   120
      ScaleHeight     =   8415
      ScaleWidth      =   6255
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      Begin zl9PacsControl.ucSplitter ucSplitter1 
         Height          =   135
         Left            =   0
         TabIndex        =   1
         Top             =   3015
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   238
         MousePointer    =   7
         SplitType       =   0
         SplitLevel      =   3
         Con1MinSize     =   1000
         Con2MinSize     =   2000
         Control1Name    =   "trvWordTree"
         Control2Name    =   "vsWordContext"
      End
      Begin VSFlex8Ctl.VSFlexGrid vsWordContext 
         Height          =   5265
         Left            =   0
         TabIndex        =   2
         Top             =   3150
         Width           =   6255
         _cx             =   11033
         _cy             =   9287
         Appearance      =   0
         BorderStyle     =   0
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   16777215
         BackColorAlternate=   16777215
         GridColor       =   14737632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   0
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   4
         SelectionMode   =   0
         GridLines       =   4
         GridLinesFixed  =   0
         GridLineWidth   =   0
         Rows            =   0
         Cols            =   2
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   -1  'True
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
         FillStyle       =   1
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   2
         WordWrap        =   -1  'True
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
         AllowUserFreezing=   3
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
         Begin RichTextLib.RichTextBox txtWordEdit 
            Height          =   1935
            Left            =   1320
            TabIndex        =   4
            Top             =   2040
            Visible         =   0   'False
            Width           =   3855
            _ExtentX        =   6800
            _ExtentY        =   3413
            _Version        =   393217
            ScrollBars      =   2
            Appearance      =   0
            AutoVerbMenu    =   -1  'True
            TextRTF         =   $"ucReportSegment.ctx":0DF4
         End
      End
      Begin MSComctlLib.TreeView trvWordTree 
         Height          =   3015
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   5318
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   176
         LineStyle       =   1
         Style           =   7
         Appearance      =   0
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
   End
   Begin VB.Image imgAdvi 
      Height          =   360
      Left            =   5400
      Picture         =   "ucReportSegment.ctx":0E91
      Top             =   8520
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgOpin 
      Height          =   360
      Left            =   4800
      Picture         =   "ucReportSegment.ctx":1593
      Top             =   8520
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgDesc 
      Height          =   360
      Left            =   4200
      Picture         =   "ucReportSegment.ctx":1C95
      Top             =   8520
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   3720
      Picture         =   "ucReportSegment.ctx":2397
      Top             =   8520
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "ucReportSegment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const MOUSEEVENTF_RIGHTDOWN = &H8 '  right button down
Private Const MOUSEEVENTF_RIGHTUP = &H10 '  right button up


Private Const MOUSEEVENTF_LEFTDOWN = &H2 '  left button down
Private Const MOUSEEVENTF_LEFTUP = &H4 '  left button up
                
Private Const NODE_BACKCOLOR_DISABLE As Long = &HF1F1F1
Private Const NODE_FORCECOLOR_DISABLE As Long = &HC0C0C0    '&H808080


Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)



Private Const LVW_KEY_WORD As String = "L"  ' 树形控件项目
Private Const LVW_KEY_NODE As String = "T"   ' 树形控件目录
  
    
Private mrsClass As ADODB.Recordset     '词句分类
Private mrsWords As ADODB.Recordset     '词句项目

Private mFileID As Long                 '报告ID
Private mstrOutLineKey As String        '词句示范内容类型,提纲关键字， “所见”“诊断” “结果”，“建议”等
Private mlngOutlineId As Long
Private mlngAdviceId As Long            '医嘱ID
Private mlngFileID As Long

Private mstrDBOwner As String              '数据库所有者
Private mintWordDblClickMode As Integer     '词句双击后的操作：0--直接写入报告；1--打开词句编辑窗口
Private mintWordPower As Integer        '词句管理权范围

Private mlngWordTreeH As Long               '词库模板树的高度
Private mlngWordShowH As Long               '词库模板内容的高度


Private mlngCurModule As Long
Private mlngCurDeptId As Long

Private mlngPatientId As Long
Private mlngPageID As Long
Private mblnAdviceMoved As Boolean

Private mblnIsInit As Boolean           '是否初始化
Private mlngExpandLevel As Long         '自动展开层级,默认为1
Private mblnIsWordValid As Boolean      '是否对词句适用条件进行判断
Private mblnAutoRemove As Boolean       '是否自动移除不可用词句及分类
Private mblnIsSyncWordFragment As Boolean

Public Event OnRequestState(ByRef lngOutlineType As TOutlineType, _
                            ByRef str所见内容 As String, ByRef str意见内容 As String, ByRef str建议内容 As String)
    
Public Event OnSendContext(ByVal strFreeText As String, _
                            ByVal str所见内容 As String, ByVal str意见内容 As String, ByVal str建议内容 As String)
                            
Public Event OnMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Property Get IsSyncWordFragment() As Boolean
    IsSyncWordFragment = mblnIsSyncWordFragment = True
End Property
                            
                            
Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property

Property Get DblWrite() As Boolean
    DblWrite = IIf(mintWordDblClickMode = 0, True, False)
End Property

Property Let DblWrite(value As Boolean)
    mintWordDblClickMode = IIf(value, 0, 1)
End Property
 

'节点数量
Property Get NodeCount() As Long
    NodeCount = trvWordTree.Nodes.Count
End Property

'选择节点类型
Property Get SelNodeType() As Long
    SelNodeType = 0
    
    If trvWordTree.SelectedItem Is Nothing Then Exit Property
    
    If Left(trvWordTree.SelectedItem.Key, 1) = LVW_KEY_WORD Then
        SelNodeType = 2
    Else
        SelNodeType = 1
    End If
End Property

'展开级别
Property Get ExpandLevel() As Long
    ExpandLevel = mlngExpandLevel
End Property

Property Let ExpandLevel(value As Long)
    mlngExpandLevel = value
    
    Call AutoExpand
End Property

'自动隐藏
Property Get AutoHide() As Boolean
    AutoHide = mblnAutoRemove
End Property

Property Let AutoHide(value As Boolean)
    mblnAutoRemove = value
    
    If mFileID <> 0 Then
        Call LoadWordClass(mFileID, mstrOutLineKey, True)
    End If
End Property


Public Sub Init(ByVal lngModuleNo As Long, ByVal lngDeptId As Long, _
    Optional ByVal blnIsForce As Boolean = False)
On Error GoTo errhandle
    If mblnIsInit And blnIsForce = False Then Exit Sub
    
    mlngCurModule = lngModuleNo
    mlngCurDeptId = lngDeptId
    
'    intWordPower=-1，不具备词句管理权;
'    intWordPower=0，全院，这时显示所有的示范，也可以更改;
'    intWordPower=1，科室，这时显示全院通用示范(科室id is null)和所在科室公有或部门内人员私有的示范，但不能更改全院通用示范;
'    intWordPower=2，个人，这时显示全院通用示范(科室id is null)和所在科室通用示范(人员id is null)和个人示范，仅个人示范可更改
    
    mintWordPower = zlGetWordPower
    
    Call InitDbOwner(glngSys)
    
    trvWordTree.ImageList = ImageList1
    
    mstrOutLineKey = ""
    
    Call InitLoaclParas
    
    mblnIsInit = True
Exit Sub
errhandle:
    mblnIsInit = False
End Sub


Public Sub SetFontSize(ByVal bytFontSize As Byte)
    FontSize = bytFontSize
    
    picBack.FontSize = bytFontSize
    vsWordContext.FontSize = bytFontSize
    Set txtWordEdit.Font = Font
    
    Set trvWordTree.Font = Font
End Sub

Private Sub InitPatientInfo(ByVal lngAdviceId As Long)
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
 
    
    strSQL = "select a.病人ID,a.主页ID, a.病人来源,a.性别,a.婴儿,b.关联id, 0 as 转储 from 病人医嘱记录 a, 影像检查记录 b Where a.id=b.医嘱id(+) and a.id=[1] " & _
        "Union All " & _
        "select a.病人ID, a.主页ID, a.病人来源,a.性别,a.婴儿,b.关联id, 1 as 转储 from H病人医嘱记录 a, H影像检查记录 b Where a.id=b.医嘱id(+) and a.id=[1] "
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "医嘱信息查询", lngAdviceId)
    If rsTemp.RecordCount <= 0 Then Exit Sub
    
    mlngPatientId = Val(nvl(rsTemp!病人ID))
    mlngPageID = Val(nvl(rsTemp!主页ID))
'    mlngPatientFrom = Val(nvl(rsTemp!病人来源))
'    mlngBabyNum = Val(nvl(rsTemp!婴儿))
    mblnAdviceMoved = IIf(Val(nvl(rsTemp!转储)) = 1, True, False)
End Sub


Private Sub LoadWordClass(FileID As Long, strOutlineKey As String, Optional blnForceRefresh As Boolean = False)
    Dim strSQL As String
    Dim rsCurClass As ADODB.Recordset
    Dim rsCurWords As ADODB.Recordset
    
    Dim rsTemp As ADODB.Recordset
    Dim objNode As Node
    Dim objPnode As Node
    Dim strKey As String
    Dim blnIsOnlyRefreshOutline As Boolean
    Dim i As Long
    Dim strUserInfo As String
    Dim strWith As String
    Dim lngIndex As Long
    Dim aryOutlineId(3) As Long
    
    
    blnIsOnlyRefreshOutline = False
    
    If FileID = mFileID And trvWordTree.Nodes.Count > 0 And blnForceRefresh = False Then
        Set rsCurClass = mrsClass
'        Set rsCurWords = mrsWords
        
        If strOutlineKey <> mstrOutLineKey Then
            '提纲不同，则对提纲进行处理
            blnIsOnlyRefreshOutline = True
        Else
            '提纲页相同时，则退出
            Exit Sub
        End If
    Else
        Set rsCurClass = Nothing
'        Set rsCurWords = Nothing
    End If
    
    mFileID = FileID
    mstrOutLineKey = strOutlineKey
    
    strSQL = "Select nvl(a.父id,0) as 提纲ID  From 病历文件结构 a" & _
             " Where a.文件ID=[1] and a.内容文本 like '%' || [2] || '%' And a.对象类型=3 And Rownum =1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询病历提纲", FileID, mstrOutLineKey)
    If rsTemp.RecordCount <= 0 Then
        trvWordTree.Nodes.Clear
        vsWordContext.Rows = 0
        txtWordEdit.Text = ""
        
        Exit Sub
    End If
    
    mlngOutlineId = Val(nvl(rsTemp!提纲id))
    
    '清空模板内容
    vsWordContext.Rows = 0
    
    If mblnAutoRemove = False Then '如果需要自动隐藏，则需要清除树节点后从新加载，如果不需隐藏，则直接设置节点状态
        If blnIsOnlyRefreshOutline Then
            Call HideOutlineNode(mlngOutlineId)
            Exit Sub
        End If
    End If
    
    '调用引用API，并且采用逆序循环删除TreeView的方法，这个方法速度更快
    Call TrvwClear
            
    If rsCurClass Is Nothing Then
        '查询词句分类
'        strSQL = "Select * from (with OutLinesTab as (" & _
'                             " Select nvl(父id,0) as 提纲ID " & _
'                             " From 病历文件结构 " & _
'                             " Where 文件ID=[1]  And 对象类型=3   ) " & _
'                        " select a.ID, a.上级ID,a.编码,a.名称,b.提纲ID " & _
'                        " from  病历词句分类 a, OutLinesTab b " & _
'                        " where  a.Id In ( " & _
'                        "                 select id " & _
'                        "                 from 病历词句分类 x " & _
'                        "                 start with x.id in( " & _
'                        "                                   select 词句分类id " & _
'                        "                                   from 病历提纲词句 a " & _
'                        "                                   Where a.提纲ID = b.提纲ID ) " & _
'                        "                 Connect By Prior 上级id=Id " & _
'                        " ) and substr(a.范围,7,1)='1') order by Id"
        strSQL = "select  distinct a.提纲ID from 病历提纲词句 a, 病历文件结构 b where a.提纲ID=nvl(父id,0) and b.文件ID=[1] and 对象类型=3"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询病历词句提纲", mlngFileID)
        
        If rsTemp.RecordCount <= 0 Then
            
            Exit Sub
        End If
        
        lngIndex = 1
        strWith = ""
        strSQL = ""
        
        
'with OutlineReleation as (select 提纲ID, 词句分类ID from 病历提纲词句 where 提纲ID=122),
'     OutlineReleation1 as (select 提纲ID, 词句分类ID from 病历提纲词句 where 提纲ID=782),
'     outlineWord as (select id,上级ID,名称,编码,范围 from 病历词句分类 where  substr(范围,7,1)='1' )
'select a.id, a.上级ID,a.名称, 提纲类别, b.提纲ID as 关联提纲 from (
'       select id,上级ID,名称,122 as 提纲类别 from outlineWord
'       start with  id in(select  词句分类ID from OutlineReleation)
'       Connect By Prior 上级id=Id
') a, OutlineReleation b
'where a.id=b.词句分类ID(+)
'
'Union All
'
'select a.id, a.上级ID,a.名称,提纲类别, b.提纲ID as 关联提纲 from (
'       select id,上级ID,名称 , 782 as 提纲类别 from outlineWord
'       start with  id in(select  词句分类ID from OutlineReleation1)
'       Connect By Prior 上级id=Id
') a, OutlineReleation1 b
'where a.id=b.词句分类ID(+)
'order by ID
    

        While Not rsTemp.EOF
            If strWith <> "" Then strWith = strWith & "," & vbCrLf
            strWith = strWith & "OutlineReleation" & lngIndex & " as (select 提纲ID, 词句分类ID from 病历提纲词句 where 提纲ID=[" & lngIndex & "])"
            
            If strSQL <> "" Then strSQL = strSQL & vbCrLf & "Union All " & vbCrLf
            strSQL = strSQL & "select a.id, a.上级ID,a.名称, a.编码, 提纲类别, b.提纲ID as 关联提纲 from (" & vbCrLf & _
                            "    select id,上级ID,名称,编码,[" & lngIndex & "] as 提纲类别 from outlineWord" & vbCrLf & _
                            "    start with  id in(select  词句分类ID from OutlineReleation" & lngIndex & ") " & vbCrLf & _
                            "    Connect By Prior 上级id=Id ) a, OutlineReleation" & lngIndex & " b where a.id=b.词句分类ID(+)"
            
            aryOutlineId(lngIndex) = Val(nvl(rsTemp!提纲id))
            lngIndex = lngIndex + 1
            Call rsTemp.MoveNext
        Wend
        
        strSQL = "select * from (with " & strWith & "," & vbCrLf & _
                        " OutlineWord as (select id,上级ID,名称,编码,范围 from 病历词句分类 where  substr(范围,7,1)='1' )" & vbCrLf & _
                        strSQL & vbCrLf & _
                        ") Order by 提纲类别, ID"
        
        
                        
        Set mrsClass = zlDatabase.OpenSQLRecord(strSQL, "查询词句分类", aryOutlineId(1), aryOutlineId(2), aryOutlineId(3))
'        Set mrsClass = zlDatabase.CopyNewRec(mrsClass)
        
        If mrsClass.RecordCount <= 0 Then Exit Sub
        
        Set rsCurClass = mrsClass
    End If
    
        '查询词句
'        strSQL = "select /*+ RULE*/ b.ID,b.分类ID,b.名称 " & _
'                   " from  病历词句分类 a,   病历词句示范 b" & IIf(mblnAutoRemove, ", Table(Cast(f_Sentence_Usable([2], [3], [4], [5]) as zlhis.t_Dic_Rowset )) C ", "") & _
'                   " where a.id=b.分类ID " & IIf(mblnAutoRemove, " and b.Id=c.编码 ", "") & " and a.Id In ( " & _
'                   "        select id " & _
'                   "        from 病历词句分类 x " & _
'                   "        start with x.id in( " & _
'                   "                select 词句分类id " & _
'                   "                from 病历提纲词句 a " & _
'                   "                where a.提纲ID in ( " & _
'                   "                       Select nvl(父id,0) as 提纲ID " & _
'                   "                       From 病历文件结构 " & _
'                   "                       Where 文件ID=[1]  And 对象类型=3 ) " & _
'                   "                     ) " & _
'                   "        Connect By Prior 上级id=Id " & _
'                   "        ) and substr(a.范围,7,1)='1' order by Id "
        strSQL = "select /*+ RULE*/ b.ID,b.分类ID,b.名称 " & _
                   " from  病历词句分类 a,   病历词句示范 b" & IIf(mblnAutoRemove, ", Table(Cast(f_Sentence_Usable([2], [3], [4], [5]) as zlhis.t_Dic_Rowset )) C ", "") & _
                   " where a.id=b.分类ID " & IIf(mblnAutoRemove, " and b.Id=c.编码 ", "") & " and a.Id In ( " & _
                   "                select 词句分类id " & _
                   "                from 病历提纲词句 a " & _
                   "                where a.提纲ID in ( " & _
                   "                       Select nvl(父id,0) as 提纲ID " & _
                   "                       From 病历文件结构 " & _
                   "                       Where 文件ID=[1]  And 对象类型=3 ) " & _
                   "                      " & _
                   "        ) and substr(a.范围,7,1)='1' order by Id "
        Set mrsWords = zlDatabase.OpenSQLRecord(strSQL, "查询词句项目", FileID, mlngOutlineId, mlngPatientId, mlngPageID, mlngAdviceId)
'        Set mrsWords = zlDatabase.CopyNewRec(mrsWords)
        
        If mrsWords.RecordCount <= 0 Then Exit Sub
        
        Set rsCurWords = mrsWords
         
     
    If mblnAutoRemove Then
        rsCurClass.Filter = "提纲类别=" & mlngOutlineId
    Else
        rsCurClass.Filter = ""
    End If
    
    rsCurClass.Sort = "编码"
    
    strUserInfo = "[" & UserInfo.用户名 & "]"
    '载入所有分类
    Do While Not rsCurClass.EOF
        
        Set objNode = Nothing
        
        On Error Resume Next
        Set objNode = trvWordTree.Nodes("T-" & rsCurClass("ID").value)
        
        If err.Number <> 0 Then
            Set objNode = Nothing
            err.Clear
        End If
        
        If zlCommFun.nvl(rsCurClass("上级id").value, 0) <> 0 Then
            Set objPnode = trvWordTree.Nodes("T-" & rsCurClass("上级id").value)
            
            If err.Number <> 0 Then
                Set objPnode = Nothing
                err.Clear
            End If
        Else
            Set objPnode = Nothing
        End If
        
        On Error GoTo errhandle
        
        If objNode Is Nothing Then
            If objPnode Is Nothing Then
                Set objNode = trvWordTree.Nodes.Add(, , "T-" & rsCurClass("ID").value, Replace(rsCurClass("名称").value, strUserInfo, ""), 2)
            Else
                Set objNode = trvWordTree.Nodes.Add("T-" & zlCommFun.nvl(rsCurClass("上级id").value, 0), tvwChild, "T-" & rsCurClass("ID").value, Replace(rsCurClass("名称").value, strUserInfo, ""), 2)
            End If
             
            objNode.tag = 0  '表示尚未加载词句节点
        End If
    
        rsCurClass.MoveNext
    Loop
    
    '隐藏不属于本提纲的词句分类节点,
    If mblnAutoRemove = False Then Call HideOutlineNode(mlngOutlineId)
    
    '根据展开层级自动展开
    Call AutoExpand
    
    Exit Sub
errhandle:
    If err.Number <> 35602 Then
        If ErrCenter() = 1 Then Resume Next
        Call SaveErrLog
    End If
End Sub

Private Sub AutoExpand()
    Dim objNode As Node
    Dim i As Long
    Dim objSelNode As Node
    
    LockWindowUpdate trvWordTree.hwnd
On Error GoTo errhandle
    Set objSelNode = trvWordTree.SelectedItem
    '根据展开层级自动展开
    For i = 1 To trvWordTree.Nodes.Count
        Set objNode = trvWordTree.Nodes(i)
        
        If Left(objNode.Key, 1) = LVW_KEY_NODE Then
            objNode.Expanded = False
            
            If GetNodeDepth(objNode) < IIf(mlngExpandLevel = 0, 999, mlngExpandLevel) Then
                objNode.Expanded = True
                
                If Val(objNode.tag) <> 1 Then
                    Call LoadWordItem(objNode)
                    objNode.tag = 1 '表示已经加载了词句节点
                End If
                
                If objNode.BackColor = NODE_BACKCOLOR_DISABLE Then objNode.Expanded = False
            End If
        End If
    Next
    
    '恢复选择的节点
    If Not objSelNode Is Nothing Then
        While GetNodeDepth(objSelNode) > IIf(mlngExpandLevel = 0, 999, mlngExpandLevel)
            Set objSelNode = objSelNode.Parent
        Wend
        
        objSelNode.Selected = True
    End If
    
    LockWindowUpdate 0
Exit Sub
errhandle:
    LockWindowUpdate 0
End Sub

Private Sub LoadWordItem(objNode As Node)
    Dim lngWordClassId As Long
    Dim objSubNode As Node
    Dim objSubClassNode As Node
    Dim i As Long
    
    lngWordClassId = Split(objNode.Key, "-")(1)
    
    mrsWords.Filter = "分类ID=" & lngWordClassId
    If mrsWords.RecordCount > 0 Then
        '加载当前节点下的词句
        Do While Not mrsWords.EOF
            On Error Resume Next
            Set objSubNode = trvWordTree.Nodes("L-" & mrsWords("ID").value)
            
            If err.Number <> 0 Then
                Set objSubNode = Nothing
                err.Clear
            End If
            
            If objSubNode Is Nothing Then
                Set objSubNode = trvWordTree.Nodes.Add(objNode, tvwChild, "L-" & mrsWords("ID").value, mrsWords("名称").value, 1)
                objSubNode.tag = -1 '表示没有进行适用性判断
            End If
            
            '判断该词句是否对该报告模板适用
            Call mrsWords.MoveNext
        Loop
    End If
    
    Set objSubClassNode = objNode.Child
    '加载子节点下的第一条词句
    While Not objSubClassNode Is Nothing
        
        lngWordClassId = Split(objSubClassNode.Key, "-")(1)
        mrsWords.Filter = "分类ID=" & lngWordClassId
        If mrsWords.RecordCount > 0 Then
            On Error Resume Next
            Set objSubNode = trvWordTree.Nodes("L-" & mrsWords("ID").value)
            
            If err.Number <> 0 Then
                Set objSubNode = Nothing
                err.Clear
            End If
        
            If objSubNode Is Nothing Then
                Set objSubNode = trvWordTree.Nodes.Add(objSubClassNode, tvwChild, "L-" & mrsWords("ID").value, mrsWords("名称").value, 1)
                objSubNode.tag = -1 '表示没有进行适用性判断
            End If
        End If
        
        Set objSubClassNode = objSubClassNode.Next
    Wend
End Sub

Private Sub HideOutlineNode(ByVal lngOutlineId As Long)
    Dim rsOutlineClass As ADODB.Recordset
    Dim i As Long
    Dim lngClassID As Long
    Dim objNode As Node
    Dim objSubNode As Node

    mrsClass.Filter = ""
    Set rsOutlineClass = mrsClass.Clone

    For i = trvWordTree.Nodes.Count To 1 Step -1
        Set objNode = trvWordTree.Nodes(i)
        
        If Left(objNode.Key, 1) = LVW_KEY_NODE Then
            lngClassID = Val(Split(objNode.Key & "-", "-")(1))
            
            rsOutlineClass.Filter = "关联提纲=" & lngOutlineId & " and ID=" & lngClassID
    
            'Node.tag:0_1_文本内容 对应说明 医嘱ID_适用状态_文本内容
            If rsOutlineClass.RecordCount > 0 Then
                '提纲存在对应分类
                objNode.BackColor = vbWhite
                objNode.ForeColor = vbBlack
            Else
                objNode.BackColor = NODE_BACKCOLOR_DISABLE
                objNode.ForeColor = NODE_FORCECOLOR_DISABLE
            End If
            
            
            If objNode.Children > 0 Then
                 Set objSubNode = objNode.Child
                 
                 While Not objSubNode Is Nothing
                     If Left(objSubNode.Key, 1) = LVW_KEY_WORD Then
                         objSubNode.BackColor = objNode.BackColor
                         objSubNode.ForeColor = objNode.ForeColor
                     End If
                     
                     Set objSubNode = objSubNode.Next
                 Wend
             End If
        End If
    Next
End Sub


Private Function GetNodeDepth(objNode As Object) As Long
'获取节点深度
    GetNodeDepth = UBound(Split(objNode.FullPath, trvWordTree.PathSeparator))
End Function


Private Sub Form_Unload(Cancel As Integer)
'    Dim strRegPath As String
'
'
'    strRegPath = "公共模块\" & App.ProductName & "\frmReportWord"
'
'    '保存词句示范区域的高度
'    '285是Pane的标题高度，使用了标题，就需要加回这个高度
'    If Not (picWordTree.Height = 0 And picWordShow.Height = 0 And picPrivateWord.Height = 0) Then
'      SaveSetting "ZLSOFT", strRegPath, "WordTreeH", picWordTree.Height
'      SaveSetting "ZLSOFT", strRegPath, "WordShowH", picWordShow.Height
'      SaveSetting "ZLSOFT", strRegPath, "PrivateWordH", picPrivateWord.Height ' + 285
'    End If
'    SaveSetting "ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\frmReportWord", "直接编辑", CLng(chk直接编辑.value)
'
'    If mblnShowWord = False Then    '通过双击打开，则显示确定和取消按钮,记录这个高度
'        SaveSetting "ZLSOFT", strRegPath, "ButtonH", picCommandButton.Height
'    End If
'
'    '保存词句示范区域的宽度
'    If mblnSingleWindow = True Then
'        strRegPath = "公共模块\" & App.ProductName & "\frmReport\SingleWindow"
'    Else
'        strRegPath = "公共模块\" & App.ProductName & "\frmReport"
'    End If
'    SaveSetting "ZLSOFT", strRegPath, "CX1", picWordTree.Width
'
'    '窗口模式,此模式下记录窗口位置
'    If mblnShowWord = False Then
'        Call SaveWinState(Me, App.ProductName)
'    End If
End Sub
 



 

'Private Sub menuAutoHide_Click()
'On Error GoTo errHandle
'    menuAutoHide.Checked = Not menuAutoHide.Checked
'    mblnAutoRemove = menuAutoHide.Checked
'
'    SaveSetting "ZLSOFT", mstrRegPrivatePath, "自动隐藏", mblnAutoRemove
'
'    Call LoadWordClass(mFileID, mstrOutLineName, True)
'Exit Sub
'errHandle:
'    MsgBoxH hWnd, err.Description, vbOKOnly, "提示"
'End Sub

Private Function GetRootHwnd() As Long
    GetRootHwnd = GetAncestor(hwnd, GA_ROOT)
End Function

Public Sub DirectWrite()
'直接写入词句
On Error GoTo errhandle
    Dim objSelNode As Node

    Set objSelNode = trvWordTree.SelectedItem
    
    If Not objSelNode Is Nothing Then
        If Left(objSelNode.Key, 1) = LVW_KEY_WORD Then
            '词句双击后，打开词句编辑窗口
            Call WriteWordDirect
        End If
    End If
Exit Sub
errhandle:
    MsgboxH GetRootHwnd, err.Description, vbOKOnly, "提示"
End Sub


Public Sub EditWrite()
On Error GoTo errhandle
    Dim objSelNode As Node

    Set objSelNode = trvWordTree.SelectedItem
    
    If Not objSelNode Is Nothing Then
        If Left(objSelNode.Key, 1) = LVW_KEY_WORD Then
            '词句双击后，打开词句编辑窗口
            WriteWordEdit Val(Split(objSelNode.Key & "-", "-")(1))
        End If
    End If
Exit Sub
errhandle:
    MsgboxH GetRootHwnd, err.Description, vbOKOnly, "提示"
End Sub


Private Sub WriteWordDirect()
'直接写入
    Dim i As Long
    Dim objNode As Node
    Dim lngApply As Long
    
    Set objNode = trvWordTree.SelectedItem
    
    If objNode Is Nothing Then Exit Sub
    If vsWordContext.Rows <= 0 Then Exit Sub
    
    lngApply = Val(Split(objNode.tag & "__", "_")(1))
    
    If lngApply = 0 Then
        If MsgboxH(GetRootHwnd, "该词句不适用于当前提纲，是否继续？", vbYesNo + vbDefaultButton2, "提示") = vbNo Then Exit Sub
    End If
    
    For i = 0 To vsWordContext.Rows - 1
        If vsWordContext.RowData(i) <> "WARING" Then
            Call DoWritWord(i, False)
        End If
    Next
End Sub
 

 

'暂时不提供分类相关处理功能
'Private Sub menuNewClass_Click()
'On Error GoTo errHandle
'    Call NewClass
'Exit Sub
'errHandle:
'    MsgBoxH GetRootHwnd,  err.Description, vbOKOnly, "提示"
'End Sub

'Private Sub NewClass()
''新增分类
'    Dim objPNode As Node
'    Dim objSubNode As Node
'    Dim strSql As String
'    Dim rsData As ADODB.Recordset
'    Dim lngPId As Long
'    Dim strPCode As String
'    Dim rsClass As ADODB.Recordset
'    Dim lngCurClassId As Long
'    Dim strCurClassCode As String
'    Dim strCurClassName As String
'    Dim i As Long
'On Error GoTo errHandle
'    Set objPNode = trvWordTree.SelectedItem
'
'    If objPNode Is Nothing Then Exit Sub
'    If Left(objPNode.Key, 1) = LVW_KEY_WORD Then Exit Sub
'
'    lngPId = Split(objPNode.Key, "-")(1)
'
'    Set rsClass = mrsClass.Clone
'    rsClass.Filter = "ID=" & lngPId
'
'    strPCode = ""
'    If rsClass.RecordCount > 0 Then strPCode = nvl(rsClass!编码)
'
'    strSql = "select 病历词句分类_ID.NEXTVAL as 词句ID from dual"
'    Set rsData = zlDatabase.OpenSQLRecord(strSql, "查询词句分类ID")
'    If rsData.RecordCount <= 0 Then
'        MsgBoxH GetRootHwnd,  "不能获取词句分类ID.", vbOKOnly, "提示"
'        Exit Sub
'    End If
'
'    lngCurClassId = Val(nvl(rsData!词句Id))
'
'
'    strSql = "select nvl(max(编码), 0) as 编码 from 病历词句分类 where 上级ID=[1]"
'    Set rsData = zlDatabase.OpenSQLRecord(strSql, "查询新增词句分类编码", lngPId)
'    If rsData.RecordCount <= 0 Then
'        MsgBoxH GetRootHwnd,  "不能获取词句分类编码.", vbOKOnly, "提示"
'        Exit Sub
'    End If
'
'    strCurClassName = "新分类1"
'    If Val(nvl(rsData!编码)) = 0 Then
'        strCurClassCode = strPCode & "01"
'    Else
'        strCurClassCode = Val(nvl(rsData!编码)) + 1
'
'        rsClass.Filter = "上级ID=" & lngPId & " and 名称='" & strCurClassName & "[" & UserInfo.用户名 & "]" & strCurClassName & "'"
'
'        i = 1
'        While rsClass.RecordCount > 0
'            i = i + 1
'            strCurClassName = "新分类" & i
'            rsClass.Filter = "上级ID=" & lngPId & " and 名称='" & "[" & UserInfo.用户名 & "]" & strCurClassName & "'"
'        Wend
'    End If
'
'    If Len(strCurClassCode) > 8 Then
'        MsgBoxH GetRootHwnd,  "分类层级已超出限制，不能继续创建子分类。", vbOKOnly, "提示"
'        Exit Sub
'    End If
'
'    strSql = "Zl_病历词句分类_Edit(1," & lngCurClassId & "," & lngPId & ",'" & strCurClassCode & "','" & _
'                                "[" & UserInfo.用户名 & "]" & strCurClassName & "','','00000010')"
'    Call zlDatabase.ExecuteProcedure(strSql, "新增词句分类")
'
'    mrsClass.AddNew
'    mrsClass!ID = lngCurClassId
'    mrsClass!上级ID = lngPId
'    mrsClass!编码 = strCurClassCode
'    mrsClass!名称 = "[" & UserInfo.用户名 & "]" & strCurClassName
'    mrsClass!提纲ID = 0
'
'    mrsClass.Update
'
'    Set objSubNode = trvWordTree.Nodes.Add(objPNode, tvwChild, "T-" & lngCurClassId, strCurClassName, 2)
'
'    objSubNode.Selected = True
'    trvWordTree.StartLabelEdit
'Exit Sub
'errHandle:
'    If ErrCenter() = 1 Then Resume
'    Call SaveErrLog
'End Sub
 
Public Function WordNew() As Boolean
    Dim strErr As String
On Error GoTo errhandle
    Dim lngCurOutline As TOutlineType
    
    Dim str所见 As String
    Dim str诊断 As String
    Dim str建议 As String

'    str所见 = "测试所见内容"
'    lngCurOutline = otDesc

    RaiseEvent OnRequestState(lngCurOutline, str所见, str诊断, str建议)

    Select Case lngCurOutline
        Case otDesc '所见
            str诊断 = ""
            str建议 = ""
        Case otOpin '诊断
            str所见 = ""
            str建议 = ""
        Case otAdvi '建议
            str所见 = ""
            str诊断 = ""
    End Select

    WordNew = WordInsert(str所见, str诊断, str建议)
Exit Function
errhandle:
    strErr = err.Description
    MsgboxH GetRootHwnd, strErr, vbOKOnly, "提示"
End Function

Public Function FullSave() As Boolean
'全套存入
    Dim strErr As String
On Error GoTo errhandle
    Dim lngCurOutline As Long
    Dim str所见 As String
    Dim str诊断 As String
    Dim str建议 As String
    
'    str所见 = "测试所见内容"
'    str诊断 = "测试诊断内容"
'    str建议 = "测试建议内容"
    
    RaiseEvent OnRequestState(lngCurOutline, str所见, str诊断, str建议)
    
    FullSave = WordInsert(str所见, str诊断, str建议)
Exit Function
errhandle:
    strErr = err.Description
    MsgboxH GetRootHwnd, strErr, vbOKOnly, "提示"
End Function


Public Function WordInsert(ByVal str所见 As String, ByVal str诊断 As String, ByVal str建议 As String) As Boolean
'插入词句
    Dim objReportWordList As New frmReportWordList
    Dim strWordContext As String
    Dim objNode As Node
    Dim objWordNode As Node
    Dim lngClassID As Long
    Dim strClassName As String
    Dim lngNewWordId As Long
    Dim strNewWordName As String
    
    
    WordInsert = False
    strWordContext = ""
    
    If Trim(str所见) <> "" Then
        strWordContext = strWordContext & "<<所见>>" & str所见
    End If
    
    If Trim(str诊断) <> "" Then
        If Trim(strWordContext) <> "" Then strWordContext = strWordContext & vbCrLf
        strWordContext = strWordContext & "<<诊断>>" & str诊断
    End If
    
    If Trim(str建议) <> "" Then
        If Trim(strWordContext) <> "" Then strWordContext = strWordContext & vbCrLf
        strWordContext = strWordContext & "<<建议>>" & str建议
    End If
                    
    Set objNode = trvWordTree.SelectedItem
    
    If Left(objNode.Key, 1) = LVW_KEY_WORD Then Set objNode = objNode.Parent
    
    lngClassID = Split(objNode.Key, "-")(1)
    strClassName = objNode.Text
    
    Call objReportWordList.ZlShowMe(Me, strWordContext, mintWordPower, _
                                    lngClassID, strClassName, _
                                    mlngCurDeptId, lngNewWordId, strNewWordName)
    If lngNewWordId <= 0 Then Exit Function
    
    Set objWordNode = trvWordTree.Nodes.Add(objNode, tvwChild, "L-" & lngNewWordId, strNewWordName, 1)
    objWordNode.tag = -1 '表示没有进行适用性判断
    
    mblnIsSyncWordFragment = True
    WordInsert = True
End Function

Public Function WordDelete() As Boolean
'删除选择词句
'删除词句示范
On Error GoTo errH
    Dim objWordNode As Node
    Dim lngWordID As Long
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    WordDelete = False
    
    Set objWordNode = trvWordTree.SelectedItem
    If Left(objWordNode.Key, 1) <> LVW_KEY_WORD Then Exit Function
    
    If MsgboxH(GetRootHwnd, "确定要删除当前选择的词句吗？", vbYesNo + vbDefaultButton2, "提示") = vbNo Then Exit Function
    
    lngWordID = Val(Split(objWordNode.Key, "-")(1))
    
    '如果词句的创建人ID 不是当前用户ID 则不允许删除这个词句
    strSQL = " SELECT 1 FROM  病历词句示范 WHERE ID=[1] AND 人员ID=[2] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "判断词句创建者", lngWordID, UserInfo.ID)
    
    If rsTemp.RecordCount > 0 Then
        strSQL = "zl_病历词句示范_delete(" & lngWordID & ")"
        
        Call zlDatabase.ExecuteProcedure(strSQL, "删除词句")
    Else
        MsgboxH GetRootHwnd, "尝试删除的词句不是当前用户创建的，不允许删除。", vbOKOnly, "提示"
        Exit Function
    End If
    
    Call trvWordTree.Nodes.Remove(objWordNode.Index)
    
    mblnIsSyncWordFragment = True
    WordDelete = True
    
    Exit Function
errH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Public Function WordModify() As Boolean
'修改选择词句
    Dim objReportWordList As New frmReportWordList
    Dim strWordContext As String
    Dim objPnode As Node
    Dim objWordNode As Node
    Dim lngClassID As Long
    Dim strClassName As String
    Dim lngWordID As Long
    Dim strWordName As String
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
    WordModify = False
    
    Set objWordNode = trvWordTree.SelectedItem
    If Left(objWordNode.Key, 1) <> LVW_KEY_WORD Then Exit Function
    
    lngWordID = Val(Split(objWordNode.Key, "-")(1))
    strWordName = objWordNode.Text
    
    '如果词句的创建人ID 不是当前用户ID 则不允许删除这个词句
    strSQL = " SELECT 1 FROM  病历词句示范 WHERE ID=[1] AND 人员ID=[2] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "判断词句创建者", lngWordID, UserInfo.ID)
    
    If rsTemp.RecordCount = 0 Then
        MsgboxH GetRootHwnd, "尝试修改的词句不是当前用户创建的，不允许修改。", vbOKOnly, "提示"
        Exit Function
    End If
         
    strWordContext = Split(objWordNode.tag & "__", "_")(2)
                    
    Set objPnode = objWordNode.Parent
    
    lngClassID = Val(Split(objPnode.Key, "-")(1))
    strClassName = objPnode.Text
    
    Call objReportWordList.ZlShowMe(Me, strWordContext, mintWordPower, _
                                    lngClassID, strClassName, _
                                    mlngCurDeptId, lngWordID, strWordName)
    If lngWordID <= 0 Then Exit Function
    
    objWordNode.tag = -1 '表示没有进行适用性判断

    Call trvWordTree_NodeClick(objWordNode)
    
    mblnIsSyncWordFragment = True
    WordModify = True
    
End Function


Private Sub trvWordTree_DblClick()
    Dim i As Integer
    Dim objSelNode As Node
    Dim strErr As String
On Error GoTo errhandle
    Set objSelNode = trvWordTree.SelectedItem
    
    If Not objSelNode Is Nothing Then
        If Left(objSelNode.Key, 1) = LVW_KEY_WORD Then
 
            If mintWordDblClickMode = 1 Then
                '词句双击后，打开词句编辑窗口
                WriteWordEdit Val(Split(objSelNode.Key & "-", "-")(1))
            Else
                Call WriteWordDirect
            End If
        End If
    End If
Exit Sub
errhandle:
    strErr = err.Description
    Call MsgboxH(GetRootHwnd, strErr, vbOKOnly, "提示")
End Sub

Private Sub LoadWordData(Node As Node)
    If Left(Node.Key, 1) = LVW_KEY_NODE Then
        If Val(Node.tag) <> 1 Then
            '载入词句项目
            Call LoadWordItem(Node)
            Node.tag = 1
            
            If mblnAutoRemove = False Then
                Call HideOutlineNode(mlngOutlineId)
            End If
        End If
    ElseIf Left(Node.Key, 1) = LVW_KEY_WORD Then
        '载入词句内容
        Call LoadWordContext(Node)
    End If
End Sub

Private Sub LoadWordContext(Node As Node)
    Dim lngWordID As Long
    Dim str内容文本 As String
    Dim aryWordLines() As TWordLine
    Dim aryPro() As String
    Dim blnIsApply As Boolean
    
On Error GoTo errhandle
    '清空原有控件
    vsWordContext.Rows = 0
    
    Call LevalEdit(False)
    
    If Left(Node.Key, 1) <> LVW_KEY_WORD Then Exit Sub
        
    ReDim aryWordLines(0)
    
    'Node.tag:0_1_文本内容 对应说明 医嘱ID_适用状态_文本内容
    aryPro = Split(Node.tag & "__", "_")
    lngWordID = Right(Node.Key, Len(Node.Key) - 2)
    
    If Val(aryPro(0)) < 0 Then
        '获取词句内容
        str内容文本 = GetWordContext(lngWordID)
    Else
        str内容文本 = aryPro(2)
    End If
    
    '解析词句内容
    Call FormatWords(lngWordID, str内容文本, aryWordLines())
    
    '判断词句适用条件
    If mblnIsWordValid Then
        If Val(aryPro(0)) <> mlngAdviceId Then
            '从数据库判断词句是否适用该检查患者报告
            blnIsApply = WordApplyState(lngWordID)
            
            If blnIsApply = False Then
                Node.BackColor = NODE_BACKCOLOR_DISABLE
                Node.ForeColor = NODE_FORCECOLOR_DISABLE
            Else
                Node.BackColor = vbWhite
                Node.ForeColor = vbBlack
            End If
        Else
            blnIsApply = IIf(Val(aryPro(1)) = 1, True, False)
        End If
        
        Node.tag = mlngAdviceId & "_" & IIf(blnIsApply, 1, 0) & "_" & str内容文本
    Else
        blnIsApply = IIf(Node.Parent.BackColor = NODE_BACKCOLOR_DISABLE, False, True)
        
        Node.tag = "0_1_" & str内容文本
    End If
    
    If blnIsApply = False Then
        '不适用处理
        vsWordContext.Rows = 1
        vsWordContext.Cell(flexcpText, 0, 1) = "注:该词句不适用此报告提纲..."
        vsWordContext.RowData(0) = "WARING"
        
        vsWordContext.BackColor = &HE0E0E0
        vsWordContext.BackColorBkg = &HE0E0E0
        
        vsWordContext.Cell(flexcpBackColor, 0, 0, 0, 1) = vbYellow
        vsWordContext.Cell(flexcpData, 0, 1) = 0
    Else
        vsWordContext.BackColor = vbWhite
        vsWordContext.BackColorBkg = vbWhite
    End If
    
    '显示词句内容
    Call ShowWordContext(aryWordLines(), blnIsApply)
    
    Exit Sub
errhandle:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function WordApplyState(ByVal lngWordID As Long) As Boolean
    '判断词句适用状态
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    
On Error GoTo errhandle
    'mblnAutoRemove如果为true，说明不适用的词句已经被自动移除，不需要后续判断
    If mblnAutoRemove Then
        WordApplyState = True
        Exit Function
    End If
    
    strSQL = "Select 编码 " & _
                " From Table(Cast(f_Sentence_Usable([1], [2], [3], [4]) as zlhis.t_Dic_Rowset )) U " & _
                " Where 编码=[5]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询词句适用状态", mlngOutlineId, mlngPatientId, mlngPageID, mlngAdviceId, lngWordID)
    
    WordApplyState = IIf(rsTemp.RecordCount <= 0, False, True)
Exit Function
errhandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetWordContext(ByVal lngWordID As Long) As String
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim str内容文本 As String
    Dim strLineText As String
    
On Error GoTo errhandle
    GetWordContext = ""
    
    strSQL = "Select 词句id,排列次序,内容性质,内容文本,诊治要素ID,替换域,要素名称,要素类型,要素长度,要素小数," & _
             " 要素单位,要素表示,要素值域,输入形态 From 病历词句组成 Where 词句ID=[1] order by 排列次序 "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询词句组成", lngWordID)
    
    If rsTemp.RecordCount <= 0 Then Exit Function
    
    str内容文本 = ""
    '从数据库中读取词句后，逐行分析并显示
    While rsTemp.EOF = False
        '先把记录中的词句内容读取到str内容文本中
        strLineText = nvl(rsTemp!内容文本)

        If rsTemp!内容性质 = 0 Then     '是自由文本，直接加入内容
            If Trim(strLineText) <> "" Then  '内容文本不为空，则解析并显示内容文本
                str内容文本 = str内容文本 & strLineText
            End If
        Else        'rsTemp!内容性质<>0 ,是要素，需要解析
            Select Case Val(nvl(rsTemp!要素表示))
                Case 0 ''文本要素解析成空“ ”
                    str内容文本 = str内容文本 & "  " & nvl(rsTemp!要素单位)
                
                Case 1 '上下
                '目前没有使用这个方式
                
                Case 2 '单选
                    str内容文本 = str内容文本 & "{{" & nvl(rsTemp!要素值域) & "}}" & nvl(rsTemp!要素单位)
                
                Case 3 '复选
                    str内容文本 = str内容文本 & "{<" & nvl(rsTemp!要素值域) & ">}" & nvl(rsTemp!要素单位)
            
            End Select
        End If
      
        rsTemp.MoveNext
    Wend
    
    GetWordContext = str内容文本
Exit Function
errhandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function GetLayoutStr() As String
'返回格式字符串[Key=picturebox1.width:20;picturebox1.height:30;]
    GetLayoutStr = "[KEY=SEGMENT@" & _
                                        GetProFmt("TRVWORDTREE.HEIGHT", trvWordTree.Height) & _
                                        GetProFmt("VSWORDCONTEXT.HEIGHT", vsWordContext.Height) & _
                                 "]"
End Function

Public Sub SetLayout(ByVal strLayout As String)
    Dim strPros As String
    Dim lngKeyIndex As String
    Dim strPro As String
    
    If Len(strLayout) <= 0 Then Exit Sub
    
    strPros = GetPros(strLayout, "SEGMENT")
    
    strPro = GetProValue(strPros, "TRVWORDTREE.HEIGHT")
    If Val(strPro) > 0 Then trvWordTree.Height = Val(strPro)
    
    strPro = GetProValue(strPro, "VSWORDCONTEXT.HEIGHT")
    If Val(strPro) > 0 Then vsWordContext.Height = Val(strPro)
    

End Sub


Private Sub ShowWordContext(aryWordLines() As TWordLine, ByVal blnIsApply As Boolean)
    Dim i As Long
    Dim lngBaseRow As Long
    Dim lngFillRow As Long
    Dim strWordOutline As String
    
    lngBaseRow = vsWordContext.Rows
    
    vsWordContext.Rows = lngBaseRow + UBound(aryWordLines)
    For i = 1 To UBound(aryWordLines)
        If Trim(aryWordLines(i).strContext) <> "" Then
            lngFillRow = lngBaseRow + (i - 1)
            strWordOutline = aryWordLines(i).strOutlineName
            
            vsWordContext.Cell(flexcpText, lngFillRow, 1) = aryWordLines(i).strContext
            vsWordContext.RowData(lngFillRow) = aryWordLines(i).strOutlineName
            
            If Len(strWordOutline) > 0 Then
                If InStr(strWordOutline, "所见") >= 1 Then
                    Set vsWordContext.Cell(flexcpPicture, lngFillRow, 0) = imgDesc.Picture
                ElseIf InStr(strWordOutline, "意见") >= 1 Or InStr(strWordOutline, "结果") >= 1 Or (InStr(strWordOutline, "诊断") >= 1 And InStr(strWordOutline, "建议") <= 0) Then
                    '匹配意见,诊断,诊断结果，排除诊断建议相关词
                    Set vsWordContext.Cell(flexcpPicture, lngFillRow, 0) = imgOpin.Picture
                ElseIf InStr(strWordOutline, "建议") >= 1 Then
                    Set vsWordContext.Cell(flexcpPicture, lngFillRow, 0) = imgAdvi.Picture
                Else
                    Set vsWordContext.Cell(flexcpPicture, lngFillRow, 0) = Image1.Picture
                End If
            Else
                Set vsWordContext.Cell(flexcpPicture, lngFillRow, 0) = Image1.Picture
            End If
            
            If blnIsApply Then
                vsWordContext.Cell(flexcpData, lngFillRow, 1) = 1
            End If
        End If
    Next
    
    vsWordContext.ColWidth(0) = 450
    If lngBaseRow <> 0 Then
        vsWordContext.Cell(flexcpAlignment, lngBaseRow, 1, vsWordContext.Rows - 1, 1) = flexAlignLeftTop
    Else
        vsWordContext.ColAlignment(1) = flexAlignLeftTop
    End If
    
    Call vsWordContext.AutoSize(0, 1)
End Sub

Private Sub trvWordTree_Expand(ByVal Node As MSComctlLib.Node)
    Dim strErr As String
On Error GoTo errhandle
'    If Node.BackColor = NODE_BACKCOLOR_DISABLE Then
'        Node.Expanded = False
'    End If
    
    Call LoadWordData(Node)
Exit Sub
errhandle:
    strErr = err.Description
    MsgboxH GetRootHwnd, strErr, vbOKOnly, "提示"
End Sub

Private Sub trvWordTree_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strErr As String
On Error GoTo errhandle
    '处理右键弹出菜单，判断是否右键
    RaiseEvent OnMouseUp(Button, Shift, X, Y)
Exit Sub
errhandle:
    strErr = err.Description
    MsgboxH GetRootHwnd, strErr, vbOKOnly, "提示"
End Sub

Private Sub trvWordTree_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim strErr As String
On Error GoTo errhandle
    Call LoadWordData(Node)
Exit Sub
errhandle:
    strErr = err.Description
    MsgboxH GetRootHwnd, strErr, vbOKOnly, "提示"
End Sub


Private Sub TrvwClear()
     Dim X As Integer
     
     With trvWordTree
        SendMessage .hwnd, WM_SETREDRAW, 0, 0
        
        For X = .Nodes.Count To 1 Step -1
            .Nodes.Remove X
        Next X
        
        SendMessage .hwnd, WM_SETREDRAW, 1, 0
     End With
End Sub
 
 
Public Sub SyncOutline(ByVal strOutlineKey As String)
'同步提纲
    Dim i As Long
    Dim strWordOutline As String
    
    
    
    If mstrOutLineKey = strOutlineKey Then Exit Sub
    
    txtWordEdit.Visible = False
'    mstrOutLineKey = strOutlineKey
        
'    If mblnAutoRemove Then
        Call LoadWordClass(mlngFileID, strOutlineKey, False)
'    End If
    
    mstrOutLineKey = strOutlineKey
    
'    If vsWordContext.Rows <= 0 Then Exit Sub
'    If Len(mstrOutLineName) <= 0 Then Exit Sub
'
'    For i = 1 To vsWordContext.Rows - 1
'        strWordOutline = vsWordContext.RowData(i)
'
'        If Len(strWordOutline) > 0 Then
'            If InStr(strWordOutline, mstrOutLineName) <= 0 Then
'                '不匹配当前选择提纲
'                vsWordContext.Cell(flexcpBackColor, 0, 0, 0, 1) = &HE0E0E0
'            Else
'                vsWordContext.Cell(flexcpBackColor, 0, 0, 0, 1) = vbWhite
'            End If
'        End If
'    Next
    
End Sub

Public Sub Refresh(ByVal lngAdviceId As Long, ByVal lngFileId As Long, _
    Optional ByVal strOutlineName As String = "所见", _
    Optional blnForceRefresh As Boolean)

    mblnIsSyncWordFragment = False
    
    If lngAdviceId <> mlngAdviceId Then
        Call InitPatientInfo(lngAdviceId)
    End If
    
    mlngAdviceId = lngAdviceId
    mlngFileID = lngFileId
    
    If lngFileId <= 0 Then
        trvWordTree.Nodes.Clear
        vsWordContext.Rows = 0
        txtWordEdit.Text = ""
        Exit Sub
    End If

    Call LoadWordClass(lngFileId, strOutlineName, blnForceRefresh)
     
End Sub

Private Sub InitDbOwner(ByVal lngSys As Long)
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL  As String
On Error GoTo errHand
    If mstrDBOwner <> "" Then Exit Sub

    strSQL = "Select 所有者 From Zlsystems Where 编号 = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取数据库所有者", lngSys)
    
    If rsTemp.RecordCount <> 0 Then mstrDBOwner = "" & rsTemp!所有者
    rsTemp.Close
    
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub InitLoaclParas()
'    Dim strSQL As String
'    Dim rsTemp As ADODB.Recordset
'
'    On Error GoTo err
'
     
    
'    mintWordDblClickMode = Val(GetDeptPara(mlngCurDeptId, "报告词句双击操作", 0))
''
'
'    mlngWordTreeH = GetSetting("ZLSOFT", strRegPath, "WordTreeH", 200)
'    mlngWordShowH = GetSetting("ZLSOFT", strRegPath, "WordShowH", 300) - 15
'    mlngPrivateWordH = GetSetting("ZLSOFT", strRegPath, "PrivateWordH", 200) + 355
'    mlngButtonH = GetSetting("ZLSOFT", strRegPath, "ButtonH", 500) + 325
''    chk直接编辑.value = IIf(CBool(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\frmReportWord", "直接编辑", False)), 1, 0)
''    ChkAutoExpand.value = IIf(CBool(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\frmReportWord", "自动展开", False)), 1, 0)
''
'
'    Exit Sub
'err:
'    If ErrCenter() = 1 Then Resume Next
'    Call SaveErrLog
End Sub

Private Sub WriteWordEdit(lngWordID As Long)
    Dim intReportViewType As TOutlineType
    Dim str所见 As String
    Dim str诊断 As String
    Dim str建议 As String
    Dim objNode As Node
    Dim objWordEdit As New frmReportWordEdit
    
    '获取当前报告的提纲类型
    RaiseEvent OnRequestState(intReportViewType, str所见, str诊断, str建议)

    Set objNode = trvWordTree.SelectedItem
    If objNode Is Nothing Then Exit Sub

    objWordEdit.zlShowMeEx Me, mlngCurDeptId, lngWordID, Split(objNode.tag & "__", "_")(2), intReportViewType, str所见, str诊断, str建议

    RaiseEvent OnSendContext("", str所见, str诊断, str建议)
End Sub

Private Sub EnterEdit(ByVal lngRow As Long, ByVal lngCol As Long)
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngWidth As Long
    Dim lngHeight As Long
    Dim lngScrollWidth As Long
    Dim lngStartSel As Long
    Dim strSelContext As String

    txtWordEdit.Visible = False
    txtWordEdit.Text = ""
    
    vsWordContext.Row = lngRow
    vsWordContext.Col = lngCol
    
    vsWordContext.EditCell
    vsWordContext.EditSelStart = 0
    vsWordContext.EditSelLength = 0
    
    mouse_event MOUSEEVENTF_LEFTDOWN Or MOUSEEVENTF_LEFTUP, 0&, 0&, 0&, 0&
    
    DoEvents    '只有执行了该句后，EditSelStart才会响应

    lngStartSel = vsWordContext.EditSelStart
    
    vsWordContext.EditSelStart = 0
    vsWordContext.EditSelLength = lngStartSel
    
    strSelContext = vsWordContext.EditSelText
    
    vsWordContext.EditSelLength = 0
    
    lngStartSel = Len(strSelContext)
    

    lngLeft = vsWordContext.ColPos(lngCol)
    lngTop = vsWordContext.RowPos(lngRow)
    lngWidth = vsWordContext.CellWidth
    lngHeight = vsWordContext.CellHeight
    
    txtWordEdit.Left = lngLeft
    txtWordEdit.Top = lngTop
    
    txtWordEdit.Width = lngWidth
    txtWordEdit.Height = lngHeight
    
    txtWordEdit.Text = vsWordContext.TextMatrix(lngRow, lngCol)
    Call SetWordStyle(txtWordEdit, vsWordContext.FontSize)
    
    txtWordEdit.Visible = True
    
    txtWordEdit.tag = -1 & "#" & lngRow & "#" & lngCol
    
    txtWordEdit.SetFocus
    txtWordEdit.SelStart = lngStartSel
    
'    mouse_event MOUSEEVENTF_LEFTDOWN Or MOUSEEVENTF_LEFTUP, 0&, 0&, 0&, 0&
End Sub

Private Sub txtWordEdit_DblClick()
    Dim strErr As String
On Error GoTo errhandle
    Call richTextBoxShowElements(txtWordEdit)
Exit Sub
errhandle:
    strErr = err.Description
    MsgboxH GetRootHwnd, strErr, vbOKOnly, "提示"
End Sub

Private Sub UserControl_Initialize()
    mlngExpandLevel = 1
    mblnIsWordValid = True
    mblnAutoRemove = False
End Sub


Private Sub UserControl_Resize()
On Error Resume Next

    picBack.Move 0, 0, ScaleWidth, ScaleHeight
'    picBack.Left = 0
'    picBack.Top = 0
'    picBack.Width = UserControl.ScaleWidth
'    picBack.Height = UserControl.ScaleHeight

    Call ucSplitter1.RePaint(False)
End Sub

Public Sub Destory()
    ucSplitter1.Destory
    
    Set mrsClass = Nothing
    Set mrsWords = Nothing
End Sub

Private Sub UserControl_Terminate()
    Call Destory
End Sub

Private Sub vsWordContext_BeforeScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long, Cancel As Boolean)
    Dim strErr As String
On Error GoTo errhandle
    If vsWordContext.Rows <= 0 Then Exit Sub
    
    If txtWordEdit.Visible Then Call LevalEdit
Exit Sub
errhandle:
    strErr = err.Description
    MsgboxH GetRootHwnd, strErr, vbOKOnly, "提示"
End Sub

Private Sub vsWordContext_Click()
    Dim strErr As String
On Error GoTo errhandle
    If vsWordContext.Rows <= 0 Then Exit Sub
    If vsWordContext.MouseRow < 0 Then
        Call LevalEdit
        Exit Sub
    End If
    
    If vsWordContext.Row < 0 Then Exit Sub
    
    If vsWordContext.Row <> vsWordContext.MouseRow Then vsWordContext.Row = vsWordContext.MouseRow
    
    If vsWordContext.Col > 0 Then
        '进入词句编辑
        If txtWordEdit.Visible Then
            Call LevalEdit
        End If
 
        If vsWordContext.Cell(flexcpBackColor, vsWordContext.Row) = vbYellow Then Exit Sub
        Call EnterEdit(vsWordContext.Row, vsWordContext.Col)
   
        Exit Sub
    ElseIf vsWordContext.Col = 0 Then
        '写入选择的词句到报告
        If txtWordEdit.Visible Then
            Call LevalEdit
        End If
        
        Call DoWritWord(vsWordContext.Row)
    End If
Exit Sub
errhandle:
    strErr = err.Description
    MsgboxH GetRootHwnd, strErr, vbOKOnly, "提示"
End Sub

Private Sub LevalEdit(Optional ByVal blnUpdateEdit As Boolean = True)
    Dim aryEditPro() As String
    
    If txtWordEdit.Visible = False Then Exit Sub
    
    If Val(txtWordEdit.tag) = -1 And blnUpdateEdit Then
        aryEditPro = Split(txtWordEdit.tag, "#")
        vsWordContext.Cell(flexcpText, Val(aryEditPro(1)), Val(aryEditPro(2))) = txtWordEdit.Text
        
        Call vsWordContext.AutoSize(0, 1)
    End If
    
    txtWordEdit.tag = ""
    txtWordEdit.Visible = False
End Sub


Private Sub vsWordContext_DblClick()
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    Dim strApplys As String
    
    If vsWordContext.Row <> 0 Then Exit Sub
    
    If vsWordContext.RowData(0) = "WARING" Then
        '获取适用条件配置
        strSQL = "select 词句ID,条件项,条件值 from 病历词句条件 Where 词句ID=[1]"
        Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询词句适用条件", Val(Split(trvWordTree.SelectedItem.Key, "-")(1)))
        
        If rsData.RecordCount <= 0 Then
            MsgboxH GetRootHwnd, "提纲词句未创建关联。", vbOKOnly, "适用条件"
            Exit Sub
        End If
        
        strApplys = ""
        While Not rsData.EOF
            strApplys = nvl(rsData!条件项) & ":" & nvl(rsData!条件值) & vbCrLf & strApplys
            rsData.MoveNext
        Wend
        
        MsgboxH GetRootHwnd, strApplys & vbCrLf & "其他可能原因：提纲词句未创建关联", vbOKOnly, "适用条件"
    End If
End Sub

Private Sub vsWordContext_KeyPress(KeyAscii As Integer)
    Dim strErr As String
On Error GoTo errhandle
    
    If KeyAscii <> 13 Then Exit Sub
    
    If vsWordContext.Rows <= 0 Then Exit Sub
    If vsWordContext.Col <> 0 Then Exit Sub
    If vsWordContext.Row < 0 Then Exit Sub
    
    Call DoWritWord(vsWordContext.Row)
Exit Sub
errhandle:
    strErr = err.Description
    MsgboxH GetRootHwnd, strErr, vbOKOnly, "提示"
End Sub

Private Sub DoWritWord(ByVal lngRow As Long, Optional ByVal blnApplyHint As Boolean = True)
    Dim strOutline As String
    Dim str所见 As String
    Dim str诊断 As String
    Dim str建议 As String
    Dim strFree As String
    
    If vsWordContext.RowData(lngRow) = "WARING" Then Exit Sub
    
    If Val(vsWordContext.Cell(flexcpData, lngRow, 1)) <> 1 And blnApplyHint Then
        If MsgboxH(GetRootHwnd, "该词句不适用于当前提纲，是否继续？", vbYesNo + vbDefaultButton2, "提示") = vbNo Then Exit Sub
    End If
    
    strOutline = vsWordContext.RowData(lngRow)
    
    If strOutline = "" Then
        strFree = vsWordContext.Cell(flexcpText, lngRow, 1)
    Else
        Select Case strOutline
            Case "<<所见>>"
                str所见 = vsWordContext.Cell(flexcpText, lngRow, 1)
            Case "<<诊断>>"
                str诊断 = vsWordContext.Cell(flexcpText, lngRow, 1)
            Case "<<建议>>"
                str建议 = vsWordContext.Cell(flexcpText, lngRow, 1)
        End Select
    End If
    
    RaiseEvent OnSendContext(strFree, str所见, str诊断, str建议)
End Sub

Private Sub vsWordContext_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errhandle
    If Button = 2 Then Exit Sub
    If vsWordContext.Rows <= 0 Then Exit Sub
    
    If vsWordContext.MouseCol <> 0 Then Exit Sub
    If vsWordContext.MouseRow < 0 Then Exit Sub
    
    If vsWordContext.Cell(flexcpPicture, vsWordContext.MouseRow, 0) Is Nothing Then Exit Sub
    If Val(vsWordContext.Cell(flexcpData, vsWordContext.MouseRow, 1)) <> 1 Then Exit Sub
    
    vsWordContext.Cell(flexcpBackColor, vsWordContext.MouseRow, 0, vsWordContext.MouseRow, 1) = &HC0FFFF
Exit Sub
errhandle:

End Sub

Private Sub vsWordContext_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errhandle
    If Button = 2 Then Exit Sub
    If vsWordContext.Rows <= 0 Then Exit Sub
    
    If vsWordContext.Col <> 0 Then Exit Sub
    If vsWordContext.Row < 0 Then Exit Sub
    
    If vsWordContext.Cell(flexcpPicture, vsWordContext.Row, 0) Is Nothing Then Exit Sub
    If Val(vsWordContext.Cell(flexcpData, vsWordContext.Row, 1)) <> 1 Then Exit Sub
    
    vsWordContext.Cell(flexcpBackColor, vsWordContext.Row, 0, vsWordContext.Row, 1) = vbWhite
Exit Sub
errhandle:

End Sub
