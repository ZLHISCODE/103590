VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{853AAF94-E49C-11D0-A303-0040C711066C}#4.3#0"; "DicomObjects.ocx"
Object = "*\Azl9PacsControl\zl9PacsControl.vbp"
Begin VB.UserControl ucReportHistory 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   9720
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6300
   ScaleHeight     =   9720
   ScaleWidth      =   6300
   Begin VB.PictureBox picBack 
      BorderStyle     =   0  'None
      Height          =   9015
      Left            =   120
      ScaleHeight     =   9015
      ScaleWidth      =   6015
      TabIndex        =   0
      Top             =   120
      Width           =   6015
      Begin zl9PacsControl.ucSplitter ucSplitter1 
         Height          =   135
         Left            =   0
         TabIndex        =   1
         Top             =   2895
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   238
         MousePointer    =   7
         SplitType       =   0
         SplitLevel      =   3
         Con1MinSize     =   1000
         Con2MinSize     =   2000
         Control1Name    =   "vsfStudy"
         Control2Name    =   "picContext"
      End
      Begin VB.PictureBox picContext 
         BorderStyle     =   0  'None
         Height          =   5985
         Left            =   0
         ScaleHeight     =   5985
         ScaleWidth      =   6015
         TabIndex        =   3
         Top             =   3030
         Width           =   6015
         Begin VB.CheckBox chkLinkView 
            BackColor       =   &H00FFFFFF&
            Caption         =   "□ 联动查看"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   375
            Left            =   4680
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   30
            Width           =   1335
         End
         Begin VB.CommandButton cmdWrite 
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   495
            Left            =   480
            Picture         =   "ucReportHistory.ctx":0000
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "写入报告内容"
            Top             =   0
            Width           =   495
         End
         Begin DicomObjects.DicomViewer dcmReportImg 
            Height          =   3495
            Left            =   240
            TabIndex        =   6
            Top             =   720
            Visible         =   0   'False
            Width           =   4455
            _Version        =   262147
            _ExtentX        =   7858
            _ExtentY        =   6165
            _StockProps     =   35
            BackColor       =   0
         End
         Begin VB.CheckBox chkReportType 
            BackColor       =   &H00FFFFFF&
            DownPicture     =   "ucReportHistory.ctx":1A72
            Height          =   495
            Left            =   0
            Picture         =   "ucReportHistory.ctx":34E4
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "文本内容"
            Top             =   0
            Width           =   495
         End
         Begin RichTextLib.RichTextBox rtxtReport 
            Height          =   5505
            Left            =   0
            TabIndex        =   4
            Top             =   480
            Width           =   6015
            _ExtentX        =   10610
            _ExtentY        =   9710
            _Version        =   393217
            ScrollBars      =   2
            Appearance      =   0
            AutoVerbMenu    =   -1  'True
            TextRTF         =   $"ucReportHistory.ctx":4F56
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfStudy 
         Height          =   2895
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   6015
         _cx             =   10610
         _cy             =   5106
         Appearance      =   1
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
         BackColorFixed  =   14737632
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16761024
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
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
      End
   End
End
Attribute VB_Name = "ucReportHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const M_STR_LISTVIEWKEY_DESCRIBE As String = "describe" '病理巨检描述标记
Private Const M_STR_LISTVIEWKET_PROCESS As String = "process"   '病理过程报告标记


Private Const M_STR_COLNAME = "序号|医嘱ID|检查号|年龄|类别|项目|部位|阴阳性|当前过程|检查时间|医嘱内容|随访描述|加载类别|关键ID|转储状态"


Private Type TPListCfg
    strSortPro As String
    strColPros As String
End Type


Public Event OnSend()
Public Event OnClick()
Public Event OnDbClick()
Public Event OnMouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event OnLinkView(ByVal lngAdviceId As Long, ByVal blnMoved As Boolean, ByVal blnIsDBClick As Boolean)

Private mblnIsInit As Boolean

Private mobjOwner As Object
Private mlngCurModule As Long
Private mlngCurDeptId As Long
Private mstrPrivs As String

Private mstrGrantDeptIds As String
Private mblnAllDepts As Boolean
Private mlngAdviceId As Long
 

Private mTPListCfg As TPListCfg

Private mstrDescriptionName As String   '检查所见段落名称
Private mstrAdviseName As String    '诊断建议段落名称
Private mstrOpinionName As String   '诊断意见段落名称
Private mstrPatholMaterialInfo As String '取材显示项目设置

Private mdtBegin As Date
Private mdtEnd As Date
Private mblnCustom As Boolean

Private mlngPatientId As Long
Private mlngPatientFrom As Long
Private mlngBabyNum As Long
Private mlngLinkID As Long

Private mstrDateRange As String
Private mblnIsThisTime As Boolean   '是否本次相关
Private mblnIsOtherDept As Boolean  '是否他科检查
Private mblnIsAutoLine As Boolean   '是否自动换行

Private mblnHistoryMoved As Boolean     '历史记录是否有进行转储
Private mblnAdviceMoved As Boolean      '当前医嘱是否进行了转储

Private mlngSelImgIndex As Long
Private mftpConTag As TFtpConTag
Private mblnAllowWrite As Boolean


'是否允许写入
Property Get AllowWrite() As Boolean
    AllowWrite = mblnAllowWrite
End Property

Property Let AllowWrite(ByVal value As Boolean)
    mblnAllowWrite = value
    
    If mblnAllowWrite = False Then
        cmdWrite.Visible = False
    Else
        If vsfStudy.Rows > 1 Then cmdWrite.Visible = True
    End If
End Property

Property Get AllowLinkViewer() As Boolean
    AllowLinkViewer = chkLinkView.Visible
End Property

Property Let AllowLinkViewer(ByVal value As Boolean)
    chkLinkView.Visible = value
End Property

'是否联动查看模式
Property Get LinkViewed() As Boolean
    LinkViewed = IIf(chkLinkView.value <> 0, True, False)
End Property

Property Let LinkViewed(ByVal value As Boolean)
    chkLinkView.value = IIf(value, 1, 0)
End Property

'历史数量
Property Get HistoryCount()
    HistoryCount = vsfStudy.Rows - 1
End Property


Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property

'日期范围
Property Get DataRange() As String
    DataRange = mstrDateRange
End Property

'本次相关
Property Get IsThisTime() As Boolean
    IsThisTime = mblnIsThisTime
End Property

Property Let IsThisTime(ByVal value As Boolean)
    mblnIsThisTime = value
End Property


'他科检查
Property Get IsOtherDept() As Boolean
    IsOtherDept = mblnIsOtherDept
End Property

Property Let IsOtherDept(ByVal value As Boolean)
    mblnIsOtherDept = value
End Property


'自动换行
Property Get IsAutoLine() As Boolean
    IsAutoLine = mblnIsAutoLine
End Property

Property Let IsAutoLine(ByVal value As Boolean)
    mblnIsAutoLine = value
End Property


'授权科室IDs
Property Get GrantDeptIds() As String
    GrantDeptIds = mstrGrantDeptIds
End Property

Property Let GrantDeptIds(ByVal value As String)
    mstrGrantDeptIds = value
End Property

'选择行
Property Get SelRow() As Long
    SelRow = vsfStudy.Row
End Property

'选择医嘱ID
Property Get SelAdviceId() As Long
    Dim intCol As Integer
    
    SelAdviceId = 0
    
    intCol = vsfStudy.ColIndex("医嘱ID")
    If intCol = -1 Then Exit Property
    
    SelAdviceId = Val(vsfStudy.TextMatrix(vsfStudy.Row, intCol))
End Property


Property Get SelMoved() As Boolean
    Dim intCol As Integer
    
    SelMoved = False
    
    intCol = vsfStudy.ColIndex("转储状态")
    If intCol = -1 Then Exit Property
    
    SelMoved = IIf(Val(vsfStudy.TextMatrix(vsfStudy.Row, intCol)) <> 0, True, False)
End Property

'选择的报告文本
Property Get SelReportText()
    SelReportText = ""
    If rtxtReport.Visible = False Then Exit Property
    
    If rtxtReport.SelLength > 0 Then
        SelReportText = rtxtReport.SelText
    Else
        SelReportText = rtxtReport.Text
    End If
End Property


Public Function IsSelected(Optional ByVal lngIndex As Long = 0) As Boolean
    Dim i As Long
    
    IsSelected = False
    If lngIndex = 0 Then
        For i = 1 To dcmReportImg.Images.Count
            If dcmReportImg.Images(i).BorderColour <> IMG_BACK_BORDER_COLOR Then
                IsSelected = True
                Exit Function
            End If
        Next
    Else
        IsSelected = IIf(dcmReportImg.Images(lngIndex).BorderColour <> IMG_BACK_BORDER_COLOR, True, False)
    End If
End Function


Public Function GetSelects() As Long()
'获取选中的图像索引
'索引从1开始
    Dim i As Long
    Dim lngBound As Long
    Dim arySelIndex() As Long
    
    ReDim arySelIndex(0)
    
    If dcmReportImg.Visible = False Then Exit Function
    
    For i = 1 To dcmReportImg.Images.Count
        
        If IsSelected(i) Then
            '如果是非透明颜色,说明是被选中的图像
            lngBound = UBound(arySelIndex) + 1
            ReDim Preserve arySelIndex(lngBound)
            
            arySelIndex(lngBound) = i
        End If
    Next
    
    GetSelects = arySelIndex
End Function
 
Public Function GetImage(ByVal lngIndex As Long) As DicomImage
'获取图像
    Dim objSelImg As DicomImage
    
    Set GetImage = Nothing
    
    If lngIndex <= 0 Or lngIndex > dcmReportImg.Images.Count Then Exit Function
    
    Set objSelImg = dcmReportImg.Images(lngIndex)
    
    Set GetImage = objSelImg.SubImage(0, 0, objSelImg.SizeX, objSelImg.SizeY, 1, 1)
End Function


Public Sub SetDateRange(ByVal strDataRange As String)
    mstrDateRange = strDataRange
End Sub




Public Sub Init(ByVal lngModuleNo As Long, ByVal lngDeptId As Long, ByVal strPrivs As String, _
    Optional ByVal blnIsForce As Boolean = False)
On Error GoTo errhandle
    If mblnIsInit And blnIsForce = False Then Exit Sub
     
    
    mlngCurModule = lngModuleNo
    mlngCurDeptId = lngDeptId
    mstrPrivs = strPrivs

    mstrDescriptionName = nvl(GetDeptPara(mlngCurDeptId, "检查所见名称", "检查所见"))
    mstrAdviseName = nvl(GetDeptPara(mlngCurDeptId, "建议名称", "诊断建议"))
    mstrOpinionName = nvl(GetDeptPara(mlngCurDeptId, "诊断意见名称", "诊断意见"))
    mstrPatholMaterialInfo = ""
    
    If mlngCurModule = G_LNG_PATHSTATION_MODULE Then
        mstrPatholMaterialInfo = zlDatabase.GetPara("取材内容设置", glngSys, mlngCurModule, "1,1,1,1,1,1,1,1,1,1")
    End If
    
    If mblnIsInit = False Then
        Call LoadControlFace
    
        Call SetFontSize(gbytFontSize)
    End If
    
    mblnIsInit = True
Exit Sub
errhandle:
    mblnIsInit = False
End Sub
 

Private Function loadPatholReportList(ByVal lngAdviceId As Long) As Integer
'根据医嘱ID 加载病理过程报告数据到历史报告列表项
'返回  0 异常   其他值: 下一个可用序号
'本过程中lvHistoryList.ListItems.Add添加的关键字分为process：过程报告  describe：巨检描述
    Dim objItem As ListItem
    Dim intCount As Integer '已经用过的序号
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    On Error GoTo errH
    
    loadPatholReportList = 0
    
    '注：病理相关表没有参与数据相关转储
    
    '加载巨检描述列表项
    strSQL = "select  a.医嘱ID, a.病理医嘱ID as 关键ID, '---' as 检查号, '---' as 年龄, '' as 影像类别, '巨检描述' as 名称, b.取材时间 as 检查时间, " & _
                    " '' as 医嘱内容, '' as 随访描述, '' as 结果阳性, 6 as 执行过程, 'describe' as 加载类别, 0 as 转储状态 " & _
                  "from 病理检查信息 a,病理取材信息 b " & _
                  "where a.病理医嘱id=b.病理医嘱id " & _
                  "and b.序号= (select min(c.序号) from 病理取材信息 c where c.病理医嘱id=a.病理医嘱id and a.医嘱id=[1]) " & _
                  "and a.医嘱id=[1]"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "加载巨检描述列表项", lngAdviceId)
    
    Call LoadList(rsTemp)
    
    '加载过程报告列表项
    strSQL = "select a.医嘱ID, b.标本名称,b.Id as 关键ID, '---' as 检查号, '---' as 年龄, '' as 影像类别, b.报告类型 as 名称,b.报告日期 as 检查时间, " & _
                    " ':' || b.标本名称 as 医嘱内容, '' as 随访描述, '' as 结果阳性, 6 as 执行过程, 'process' as 加载类别, 0 as 转储状态  " & _
                  "from 病理检查信息 a ,病理过程报告 b " & _
                  "where a.病理医嘱id=b.病理医嘱id and a.医嘱id=[1] " & _
                  "order by b.报告日期 "
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "加载过程报告列表项", lngAdviceId)
    
    Call LoadList(rsTemp)
    
    loadPatholReportList = vsfStudy.Rows - 1
    
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
End Function

Private Function getReportType(ByVal strType As String) As String
'获得具体报告类型  参数：数据库中的数字
    getReportType = ""
    
    Select Case strType
        Case "0"
            getReportType = "冰冻报告"
        Case "1"
            getReportType = "免疫报告"
        Case "2"
            getReportType = "分子报告"
        Case "3"
            getReportType = "特染报告"
        Case Else
            getReportType = strType
    End Select
End Function


Private Sub InitPatientInfo()
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
 
    
    strSQL = "select a.病人ID,a.主页ID, a.病人来源,a.性别,a.婴儿,b.关联id, 0 as 转储 from 病人医嘱记录 a, 影像检查记录 b Where a.id=b.医嘱id(+) and a.id=[1] " & _
        "Union All " & _
        "select a.病人ID, a.主页ID, a.病人来源,a.性别,a.婴儿,b.关联id, 1 as 转储 from H病人医嘱记录 a, H影像检查记录 b Where a.id=b.医嘱id(+) and a.id=[1] "
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "医嘱信息查询", mlngAdviceId)
    If rsTemp.RecordCount <= 0 Then Exit Sub
    
    mlngPatientId = Val(nvl(rsTemp!病人ID))
    mlngPatientFrom = Val(nvl(rsTemp!病人来源))
    mlngBabyNum = Val(nvl(rsTemp!婴儿))
    mlngLinkID = Val(nvl(rsTemp!关联ID))
    mblnAdviceMoved = IIf(Val(nvl(rsTemp!转储)) = 1, True, False)
End Sub

Private Sub LoadNormalReportList(ByVal lngAdviceId As Long, _
    Optional ByVal dtBegin As Date = 0, _
    Optional ByVal dtEnd As Date = 0)
    
    Dim strSQL As String
    Dim strTime As String
    Dim strTemp As String
    Dim rsTemp As ADODB.Recordset
    
    If mlngCurModule <> G_LNG_PATHOLSYS_NUM Then
        strSQL = "Select A.ID 医嘱ID,A.ID as 关键ID,A.开嘱时间  开嘱时间,A.医嘱内容, B.执行过程, B.结果阳性,C.检查号, " & _
               " C.影像类别,C.随访描述,C.年龄,C.接收日期 检查时间,E.名称,E.标本部位,'' as 加载类别,0 as 转储状态 " & _
               " From 病人医嘱记录 A,病人医嘱发送 B,影像检查记录 C,病人信息 D,诊疗项目目录 E" & _
               " Where A.病人id = [1] And A.相关id Is Null And B.医嘱ID=A.ID and a.病人id = d.病人id  " & _
               " AND A.ID=C.医嘱ID AND A.诊疗项目ID = E.ID AND b.执行过程 >= 2 "
    Else
        strSQL = "Select A.ID 医嘱ID,A.ID as 关键ID,A.开嘱时间  开嘱时间,A.医嘱内容, B.执行过程, B.结果阳性,C.病理号 检查号," & _
               " F.影像类别,F.随访描述,F.年龄,C.报到时间 检查时间,E.名称,E.标本部位,'' as 加载类别,0 as 转储状态 " & _
               " From 病人医嘱记录 A,病人医嘱发送 B,影像检查记录 F,病理检查信息 C,病人信息 D,诊疗项目目录 E" & _
               " Where A.病人id = [1] And A.相关id Is Null And B.医嘱ID=A.ID and a.病人id = d.病人id " & _
               " AND A.ID=C.医嘱ID(+) AND A.诊疗项目ID = E.ID and a.id=F.医嘱ID AND b.执行过程 >= 2 "
    End If
    
    If dtBegin <> 0 And dtEnd <> 0 Then
        strSQL = strSQL & " AND B.发送时间 between [6] and [7]"
    End If
    
    '本次检查
    If mblnIsThisTime And mlngPatientFrom = 2 Then
        strSQL = strSQL & " And (A.病人来源=2 And A.主页ID=D.主页ID)"
    End If
    
    '它科检查
    If mblnAllDepts = False Then
        If Not mblnIsOtherDept Then
            strSQL = strSQL & " And A.执行科室id+0 =[2] "
        Else
            strSQL = strSQL & " And  (A.执行科室id+0 <>[2] and B.执行过程 >= 5 or A.执行科室id+0 =[2]) "
        End If
    Else
        strSQL = strSQL & " And (Instr( [3],',' || A.执行科室id || ',' ) >0)"
    End If
    
    '婴儿
    strSQL = strSQL & " And NVL(A.婴儿,0) = [8]"
    
    '启用关联病人，才查询关联ID
    If mlngLinkID <> 0 Then
        If mlngCurModule <> G_LNG_PATHOLSYS_NUM Then
            strSQL = strSQL & " union select A.ID 医嘱ID,A.ID as 关键ID,A.开嘱时间  开嘱时间,A.医嘱内容, B.执行过程, " & _
                " B.结果阳性,C.检查号, C.影像类别,C.随访描述,C.年龄,C.接收日期 检查时间,E.名称,E.标本部位,'' as 加载类别, 0 as 转储状态 " & _
                " From 病人医嘱记录 A ,病人医嘱发送 B,影像检查记录 C,病人信息 D,诊疗项目目录 E" & _
                " Where B.医嘱ID=A.ID AND A.ID=C.医嘱ID and a.病人id = d.病人id AND A.诊疗项目ID = E.ID AND b.执行过程 >= 2 AND A.id in (Select 医嘱ID from 影像检查记录 Where 关联ID =[4]) "
        Else
            strSQL = strSQL & " union select A.ID 医嘱ID,A.ID as 关键ID,A.开嘱时间  开嘱时间,A.医嘱内容, B.执行过程, " & _
                " B.结果阳性,C.病理号 检查号,F.影像类别,F.随访描述,F.年龄,C.报到时间 检查时间, E.名称,E.标本部位,'' as 加载类别, 0 as 转储状态 " & _
                " From 病人医嘱记录 A,病人医嘱发送 B,影像检查记录 F,病理检查信息 C,病人信息 D,诊疗项目目录 E" & _
                " Where A.id in (Select 医嘱ID from 影像检查记录 Where 关联ID =[4]) And B.医嘱ID=A.ID and a.id=C.医嘱ID(+) and a.病人id = d.病人id AND A.诊疗项目ID = E.ID and a.id=F.医嘱ID and b.执行过程 >= 2 "
        End If
        
        If dtBegin <> 0 And dtEnd <> 0 Then
            strSQL = strSQL & " AND B.发送时间 between [6] and [7]"
        End If
        
        '本次检查
        If mblnIsThisTime And mlngPatientFrom = 2 Then
            strSQL = strSQL & " And (A.病人来源=2 And A.主页ID=D.主页ID)"
        End If
        
'        '它科检查
'        If chkOtherDeptReport.Value <> 1 Then
'            strSql = strSql & " And c.执行科室id+0 in(select  部门id  from 部门人员 where 人员id = [5] union all select to_Number([2]) from dual) "
'        End If
        '它科检查
        If mblnAllDepts = False Then
            If Not mblnIsOtherDept Then
                strSQL = strSQL & " And A.执行科室id+0 =[2] "
            Else
                strSQL = strSQL & " And  (A.执行科室id+0 <>[2] and B.执行过程 >= 5 or A.执行科室id+0 =[2]) "
            End If
        Else
            strSQL = strSQL & " And (Instr( [3],',' || A.执行科室id || ',' ) >0)"
        End If
        
        strSQL = strSQL & " And NVL(A.婴儿,0) = [8]"
    End If
    
    If mblnHistoryMoved Then
        strTemp = Replace(strSQL, "0 as 转储状态", "1 as 转储状态")
        strTemp = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
        strTemp = Replace(strTemp, "病人医嘱发送", "H病人医嘱发送")
        strTemp = Replace(strTemp, "影像检查记录", "H影像检查记录")
        strTemp = Replace(strTemp, "病人检查信息", "H病人检查信息")
        strSQL = strSQL & vbNewLine & " Union ALL " & vbNewLine & strTemp
    End If
    
    strSQL = "Select * From (" & vbNewLine & strSQL & vbNewLine & ") Order By 开嘱时间 Asc"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "", mlngPatientId, _
            mlngCurDeptId, "," & mstrGrantDeptIds & ",", mlngLinkID, UserInfo.ID, dtBegin, dtEnd, mlngBabyNum)
    
    If rsTemp.RecordCount > 0 Then

        rsTemp.Filter = "医嘱id <> " & lngAdviceId
        
        Call LoadList(rsTemp)
        
        If mblnIsAutoLine Then
            vsfStudy.WordWrap = True
            vsfStudy.AutoSize 1, vsfStudy.Cols - 1
        Else
            vsfStudy.WordWrap = False
            vsfStudy.AutoSize 1, vsfStudy.Cols - 1
        End If
    
    End If
End Sub

Private Sub LoadList(rsTemp As ADODB.Recordset)
    Dim lngAdviceIdIndex As Long    '医嘱ID列索引
    Dim lngSeriesNoIndex As Long    '序号列索引
    Dim lngStudyNoIndex As Long     '检查号列索引
    Dim lngAgeIndex As Long         '年龄列索引
    Dim lngKindIndex As Long        '类别列索引
    Dim lngItemIndex As Long        '项目列索引
    Dim lngProcedureIndex As Long   '当前过程列索引
    Dim lngMasculineIndex As Long   '阴阳性列索引
    Dim lngFollowUpIndex As Long    '随访列索引
    Dim lngAdviceContextIndex As Long '医嘱内容
    Dim lngCheckPointIndex As Long  '检查部位
    Dim lngLoadTypeIndex As Long
    Dim lngKyeIdIndex As Long
    Dim lngMovedStateIndex As Long
    Dim lngRow As Long
    
    If rsTemp.RecordCount <= 0 Then Exit Sub
    
    With vsfStudy
        lngAdviceIdIndex = .ColIndex("医嘱ID")
        lngSeriesNoIndex = .ColIndex("序号")
        lngStudyNoIndex = .ColIndex("检查号")
        lngAgeIndex = .ColIndex("年龄")
        lngKindIndex = .ColIndex("类别")
        lngItemIndex = .ColIndex("项目")
        lngProcedureIndex = .ColIndex("当前过程")
        lngMasculineIndex = .ColIndex("阴阳性")
        lngFollowUpIndex = .ColIndex("随访描述")
        lngAdviceContextIndex = .ColIndex("医嘱内容")
        lngCheckPointIndex = .ColIndex("部位")
        lngLoadTypeIndex = .ColIndex("加载类别")
        lngKyeIdIndex = .ColIndex("关键ID")
        lngMovedStateIndex = .ColIndex("转储状态")
        
        If mlngCurModule = G_LNG_PATHOLSYS_NUM Then .ColHidden(lngKindIndex) = True
        
        Do While Not rsTemp.EOF
            .Rows = .Rows + 1
'            iCount = .Rows
            lngRow = .Rows - 1
            
            .TextMatrix(lngRow, lngAdviceIdIndex) = Val(nvl(rsTemp!医嘱ID))
            .TextMatrix(lngRow, lngKyeIdIndex) = Val(nvl(rsTemp!关键ID))
            .TextMatrix(lngRow, lngLoadTypeIndex) = nvl(rsTemp!加载类别)
            .TextMatrix(lngRow, lngMovedStateIndex) = nvl(rsTemp!转储状态)

            .TextMatrix(lngRow, lngSeriesNoIndex) = lngRow 'iCount
            .TextMatrix(lngRow, lngStudyNoIndex) = nvl(rsTemp!检查号)
            .TextMatrix(lngRow, lngAgeIndex) = nvl(rsTemp!年龄)
            
            If mlngCurModule <> G_LNG_PATHOLSYS_NUM Then
                .TextMatrix(lngRow, lngKindIndex) = nvl(rsTemp!影像类别)
                .TextMatrix(lngRow, lngItemIndex) = nvl(rsTemp!名称)
            Else
                .TextMatrix(lngRow, lngItemIndex) = getReportType(nvl(rsTemp!名称))
            End If
            
            
            
            .TextMatrix(lngRow, lngProcedureIndex) = Decode(Val(nvl(rsTemp!执行过程, 0)), -1, "已驳回", 0, "已登记", 1, "已登记", _
                                                2, "已报到", 3, "已检查", 4, "已报告", 5, "已审核", "已完成")
            .Cell(flexcpData, lngRow, lngProcedureIndex) = Val(nvl(rsTemp!执行过程, 0))
            
            .TextMatrix(lngRow, lngMasculineIndex) = IIf(Val(nvl(rsTemp!结果阳性)) = 1, "阳", "")
            
            .TextMatrix(lngRow, lngFollowUpIndex) = nvl(rsTemp!随访描述)
            .TextMatrix(lngRow, .ColIndex("检查时间")) = Format(rsTemp!检查时间, "yyyy-MM-dd hh:mm")
            
            
            If UBound(Split(nvl(rsTemp!医嘱内容), ":")) > 0 Then
                .TextMatrix(lngRow, lngAdviceContextIndex) = Split(nvl(rsTemp!医嘱内容), ":")(0)
                .TextMatrix(lngRow, lngCheckPointIndex) = Split(nvl(rsTemp!医嘱内容), ":")(1)
            Else
                .TextMatrix(lngRow, lngAdviceContextIndex) = nvl(rsTemp!医嘱内容)
                .TextMatrix(lngRow, lngCheckPointIndex) = ""
            End If
            
            rsTemp.MoveNext
'                   If .Rows > 1 Then .Row = 1
        Loop
    End With
        
End Sub

Private Sub GetDateRange(ByRef dtBegin As Date, ByRef dtEnd As Date)
    Dim blnNoTime As Boolean
    
    blnNoTime = False
    
    '获取时间范围
    Select Case mstrDateRange
        Case "一个月"
            dtBegin = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd 00:00:00")) - 30
            dtEnd = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd 23:59:59"))
        Case "二个月"
            dtBegin = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd 00:00:00")) - 60
            dtEnd = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd 23:59:59"))
        Case "三个月"
            dtBegin = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd 00:00:00")) - 90
            dtEnd = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd 23:59:59"))
        Case "半年"
            dtBegin = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd 00:00:00")) - 180
            dtEnd = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd 23:59:59"))
        Case "一年"
            dtBegin = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd 00:00:00")) - 365
            dtEnd = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd 23:59:59"))
        Case "两年"
            dtBegin = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd 00:00:00")) - 730
            dtEnd = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd 23:59:59"))
        Case "三年"
            dtBegin = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd 00:00:00")) - 1095
            dtEnd = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd 23:59:59"))
        Case "不限"
            blnNoTime = True
        Case "自定义"
            dtBegin = mdtBegin
            dtEnd = mdtEnd
    End Select
    
    '不限时间时，用dtBegin + 1（1899/12/31）去判断是否转存，不能用dtBegin或直接blnMoved = true
    If blnNoTime Then
        mblnHistoryMoved = MovedByDate(dtBegin + 1)
    Else
        mblnHistoryMoved = MovedByDate(dtBegin)
    End If
End Sub


Private Sub ResetFace()
    vsfStudy.Rows = 1
    rtxtReport.Text = ""
    rtxtReport.Visible = True
    dcmReportImg.Images.Clear
    dcmReportImg.Visible = False
    chkReportType.value = Unchecked
    
    mftpConTag.Ip = ""
End Sub


Private Function GetRootHwnd() As Long
    GetRootHwnd = GetAncestor(hwnd, GA_ROOT)
End Function

Public Function Refresh(ByVal lngAdviceId As Long, Optional ByVal blnForce As Boolean) As Boolean
'刷新历史检查列表
    Dim dtBegin As Date
    Dim dtEnd As Date
    
    On Error GoTo errhandle

    If lngAdviceId <= 0 Then
        vsfStudy.Rows = 1
        chkReportType.Enabled = False
        cmdWrite.Enabled = False
        
        Exit Function
    Else
        chkReportType.Enabled = True
'        cmdWrite.Enabled = True
    End If
    
    If mlngAdviceId = lngAdviceId And Not blnForce Then Exit Function
    
    Call ResetFace
  
    mlngAdviceId = lngAdviceId
    
    Call InitPatientInfo
     
    Call GetDateRange(dtBegin, dtEnd)
    
    If mlngCurModule = G_LNG_PATHSTATION_MODULE Then
        Call loadPatholReportList(lngAdviceId)
    End If
    
    Call LoadNormalReportList(lngAdviceId, dtBegin, dtEnd)
    

'    If mTPListCfg.strList <> "" Then
'        Call DoLoadListCfg(mTPListCfg.strList)
'    End If
    
    If mTPListCfg.strSortPro <> "" Then
        Call DoLoadListSort(mTPListCfg.strSortPro)
    End If
    
    chkReportType.Enabled = vsfStudy.Rows > 1
    cmdWrite.Enabled = vsfStudy.Row > 1
    cmdWrite.Visible = mblnAllowWrite
    
    Refresh = True
    Exit Function
errhandle:
    MsgboxH GetRootHwnd, err.Description, vbOKOnly, "提示"
    err.Clear
End Function

Public Sub ShowDateConfig()
    Dim objSetTime As New frmSetTime
    
    Call objSetTime.ShowSetTime(mdtBegin, mdtEnd, Me)
End Sub


Private Sub LoadControlFace()
On Error GoTo errhandle
    Dim strValue As String
    Dim objControl As CommandBarControl
           
 
    strValue = GetSetting("ZLSOFT", "私有模块\" & UserInfo.用户名 & "\" & App.ProductName & "\ReportHistory", "列表配置", "")
        
    If InStr(strValue, ";") > 0 Then
        mTPListCfg.strColPros = Split(strValue, ";")(1)
        mTPListCfg.strSortPro = Split(strValue, ";")(0)
    Else
        mTPListCfg.strColPros = M_STR_COLNAME
        mTPListCfg.strSortPro = ""
    End If
    
    '判断列数量是否一致，如果不一致，则使用默认列配置...

    Call GridInit(mTPListCfg.strColPros)
   
    
    mdtBegin = CDate(Format(zlDatabase.Currentdate - 365, "yyyy-mm-dd 00:00:00"))
    mdtEnd = CDate(Format(zlDatabase.Currentdate, "yyyy-mm-dd 23:59:59"))
    
    
    Exit Sub
errhandle:
    MsgboxH GetRootHwnd, err.Description, vbExclamation, "提示"
    err.Clear
End Sub

Private Sub GridInit(strColName As String)
On Error GoTo errH
    '初始化配置列表
    Dim i As Integer
    Dim lngCount As Long
    Dim arrData() As String
    
    Dim strColPros As String
    Dim aryColPro() As String
    
    arrData = Split(strColName, "|")
    lngCount = UBound(arrData) + 1

    With vsfStudy
    
        .Cols = lngCount
        .FixedRows = 1
        .FixedCols = 0
        .RowHeightMin = 320
'        .Cell(flexcpAlignment, 0, 0, 0, lngCount - 1) = flexAlignCenterCenter

        '最后一列自动填充满列表
        .AllowUserResizing = flexResizeColumns
        .ExtendLastCol = True
        .AutoResize = True
        .ExplorerBar = 7 '用于列头拖动和排序
        .AutoSizeMode = flexAutoSizeRowHeight
    
        .WordWrap = True
        .AutoSizeMouse = True
        .SelectionMode = flexSelectionByRow
        .ScrollTrack = True
        
        For i = 0 To lngCount - 1
            strColPros = arrData(i) & ",,,,,"
            aryColPro = Split(strColPros, ",")
            
            .TextMatrix(0, i) = aryColPro(0)
            .ColKey(i) = aryColPro(0)
            
            If Val(aryColPro(1)) > 0 Then
                .ColWidth(i) = Val(aryColPro(1))
            End If
            
            If CBool(Val(aryColPro(2))) = True Then
                .ColHidden(i) = True
            Else
                .ColHidden(i) = False
            End If
        Next
        
        .Rows = 1
        If .Rows > 1 Then .RowSel = 1
        
        .ColHidden(.ColIndex("医嘱ID")) = True '隐藏医嘱ID
        .ColHidden(.ColIndex("加载类别")) = True '隐藏加载类别
        .ColHidden(.ColIndex("关键ID")) = True '隐藏关键ID
        .ColHidden(.ColIndex("转储状态")) = True '隐藏关键ID
    End With
    Exit Sub
errH:
    MsgboxH GetRootHwnd, err.Description, vbExclamation, "提示"
End Sub

Public Sub SetFontSize(ByVal bytFontSize As Byte)
    Dim CtlFont As StdFont
    
    Set CtlFont = New StdFont
    CtlFont.Size = bytFontSize
    
    UserControl.FontSize = bytFontSize

    Call SetColWithd(bytFontSize)
    
    vsfStudy.FontSize = bytFontSize
     
End Sub

Private Sub SetColWithd(ByVal bytSize As Long)
    With vsfStudy
        Select Case bytSize
            Case 9
                .ColWidth(.ColIndex("序号")) = 500
                .ColWidth(.ColIndex("部位")) = 1000
                .ColWidth(.ColIndex("当前过程")) = 900
                .ColWidth(.ColIndex("检查号")) = 700
                .ColWidth(.ColIndex("类别")) = 700
                .ColWidth(.ColIndex("年龄")) = 700
                .ColWidth(.ColIndex("检查时间")) = 1600
                .ColWidth(.ColIndex("项目")) = 1200
                .ColWidth(.ColIndex("阴阳性")) = 800
            Case 12
                .ColWidth(.ColIndex("序号")) = 600
                .ColWidth(.ColIndex("部位")) = 1250
                .ColWidth(.ColIndex("当前过程")) = 1100
                .ColWidth(.ColIndex("检查号")) = 900
                .ColWidth(.ColIndex("类别")) = 900
                .ColWidth(.ColIndex("年龄")) = 900
                .ColWidth(.ColIndex("检查时间")) = 2200
                .ColWidth(.ColIndex("项目")) = 1450
                .ColWidth(.ColIndex("阴阳性")) = 1000
            Case 15
                .ColWidth(.ColIndex("序号")) = 700
                .ColWidth(.ColIndex("部位")) = 1500
                .ColWidth(.ColIndex("当前过程")) = 1300
                .ColWidth(.ColIndex("检查号")) = 1100
                .ColWidth(.ColIndex("类别")) = 1100
                .ColWidth(.ColIndex("年龄")) = 1100
                .ColWidth(.ColIndex("检查时间")) = 2800
                .ColWidth(.ColIndex("项目")) = 1700
                .ColWidth(.ColIndex("阴阳性")) = 1200
        End Select
    End With
End Sub

 

Private Sub chkLinkView_Click()
On Error GoTo errhandle
    chkLinkView.Caption = IIf(chkLinkView.value <> 0, "■ 联动查看", "□ 联动查看")
    chkLinkView.ForeColor = IIf(chkLinkView.value <> 0, vbBlack, &H808080)
    
    If chkLinkView.value = 0 Then
        RaiseEvent OnLinkView(0, False, False)
        Exit Sub
    End If
    
    Call DoLinkView(False)
    
Exit Sub
errhandle:
    MsgboxH hwnd, err.Description, vbOKOnly, "提示"
End Sub

Public Sub CloseLinkViewer()
On Error GoTo errhandle
    RaiseEvent OnLinkView(0, False, False)
Exit Sub
errhandle:
    MsgboxH hwnd, err.Description, vbOKOnly, "提示"
End Sub

Private Sub DoLinkView(ByVal blnIsDBClick As Boolean)
    Dim intCol As Long
    Dim lngAdviceId As Long
    Dim blnMoved As Boolean
On Error GoTo errhandle
    
    If vsfStudy.Rows <= 1 Then Exit Sub
    
    intCol = vsfStudy.ColIndex("医嘱ID")
    If intCol = -1 Then Exit Sub
    
    lngAdviceId = Val(vsfStudy.TextMatrix(vsfStudy.Row, intCol))
    blnMoved = IIf(Val(vsfStudy.TextMatrix(vsfStudy.Row, vsfStudy.ColIndex("转储状态"))) = 0, False, True)
    
    RaiseEvent OnLinkView(lngAdviceId, blnMoved, blnIsDBClick)
    
Exit Sub
errhandle:
    MsgboxH hwnd, err.Description, vbOKOnly, "提示"
End Sub

Private Sub chkReportType_Click()
    
On Error GoTo errhandle
    cmdWrite.Enabled = False
    
    If chkReportType.value = 1 Then
 
        '载入报告图像
        Call ViewReportImage
        
        chkReportType.ToolTipText = "报告图像"
    Else
        '载入报告描述
        Call ViewReportContext
        
        chkReportType.ToolTipText = "文本内容"
    End If
Exit Sub
errhandle:
    MsgboxH hwnd, err.Description, vbOKOnly, "提示"
End Sub


Public Sub ViewReportImage()
'查看报告图
    Dim strLoadType As String
    Dim blnMoved As Boolean
    
    If SelAdviceId <= 0 Then
        MsgboxH hwnd, "请选择需要查看的历史记录。", vbOKOnly, "提示"
        Exit Sub
    End If
    
    blnMoved = IIf(Val(vsfStudy.TextMatrix(vsfStudy.Row, vsfStudy.ColIndex("转储状态"))) = 0, False, True)
    strLoadType = vsfStudy.TextMatrix(vsfStudy.Row, vsfStudy.ColIndex("加载类别"))
    
    If Len(strLoadType) > 0 Then
        MsgboxH hwnd, "该报告类别下不能显示报告图。", vbOKOnly, "提示"
        Exit Sub
    End If
         
    If SelAdviceId > 0 And SelAdviceId <> dcmReportImg.tag Then

        Call LoadReportImage(SelAdviceId, blnMoved)
        dcmReportImg.tag = SelAdviceId
    End If
    
    dcmReportImg.Visible = True
    rtxtReport.Visible = False
    
    chkReportType.value = Checked
End Sub

Public Sub ViewReportContext()
'加载报告文本
    Dim strLoadType As String
    Dim blnMoved As Boolean
    
    If SelAdviceId <= 0 Then
        MsgboxH hwnd, "请选择需要查看的历史记录。", vbOKOnly, "提示"
        Exit Sub
    End If
    
    blnMoved = IIf(Val(vsfStudy.TextMatrix(vsfStudy.Row, vsfStudy.ColIndex("转储状态"))) = 0, False, True)
    strLoadType = vsfStudy.TextMatrix(vsfStudy.Row, vsfStudy.ColIndex("加载类别"))
    
    If SelAdviceId > 0 And SelAdviceId <> rtxtReport.tag Then
        Call LoadReport(SelAdviceId, _
            strLoadType, _
            vsfStudy.TextMatrix(vsfStudy.Row, vsfStudy.ColIndex("关键ID")), _
            blnMoved)
            
        rtxtReport.tag = SelAdviceId
    End If
    
    rtxtReport.Visible = True
    dcmReportImg.Visible = False
    
    chkReportType.value = Unchecked
End Sub

Private Sub LoadReportImage(ByVal lngAdviceId As Long, ByVal blnMoveState As Boolean)
'报告图查询...
    Dim strSQL As String
    Dim strPicSql As String
    Dim rsData As ADODB.Recordset
    Dim lngFileId As Long
    
    strSQL = "select 病历ID from 病人医嘱报告 where 医嘱ID=[1] "
    If blnMoveState Then
        strSQL = Replace(strSQL, "病人医嘱报告", "H病人医嘱报告")
    End If
    
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "历史医嘱报告ID查询", lngAdviceId)
    
    dcmReportImg.Images.Clear
    If rsData.RecordCount <= 0 Then Exit Sub
    
    
    lngFileId = Val(nvl(rsData!病历Id))
    
    '从电子病历内容中查询数据
    strSQL = "Select  Id As 表格Id From 电子病历内容" & _
                " Where 文件id = [1] And 对象类型 = 3 And Substr(对象属性, Instr(对象属性, ';', 1, 18) + 1, 1) = '2' " & _
                " Order By 对象序号"
                
    strPicSql = "select ID,文件ID,父ID,开始版,对象标记,对象属性,内容行次 from 电子病历内容 where  文件ID=[1] and 父ID=[2] and 对象类型=5 order by 对象标记"
        
    '是否转储处理
    If blnMoveState Then
        strSQL = Replace(strSQL, "电子病历内容", "H电子病历内容")
        strPicSql = Replace(strPicSql, "电子病历内容", "H电子病历内容")
    End If
    
    
    '读取报告图信息****************************************
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询报告图框", lngFileId)
    
    If rsData.RecordCount > 0 Then
        '读取标记图，报告图
        '图像对象查询
        dcmReportImg.tag = Val(nvl(rsData!表格ID))
        
        Set rsData = zlDatabase.OpenSQLRecord(strPicSql, "查询报告图片", lngFileId, Val(nvl(rsData!表格ID)))
        If rsData.RecordCount > 0 Then
            
            Call ParshReportImgData(lngAdviceId, rsData, blnMoveState)
        End If
    End If
    
  
End Sub


Private Sub ParshReportImgData(ByVal lngAdviceId As Long, rsData As ADODB.Recordset, ByVal blnMoveState As Boolean)
'解析报告图像数据
    Dim aryImgPro() As String
    Dim reportImgTag As TReportImgTag
    Dim result As ftpResult
    Dim blnIsAbort As Boolean
    Dim objDcmImg As DicomImage
    
 
    If rsData Is Nothing Then Exit Sub
    
    rsData.MoveFirst
    blnIsAbort = False
    
    While Not rsData.EOF
        '第一个属性说明：0普通图像，1标记图像，2报告图像
        aryImgPro = Split(nvl(rsData!对象属性) & ";;;;;;;;;;;;;;;;;;;;", ";")
        
        reportImgTag.lngFileId = Val(rsData!文件ID)
        reportImgTag.lngTableId = Val(rsData!父ID)
        reportImgTag.strObjectTag = Val(rsData!对象标记)
        reportImgTag.strPros = nvl(rsData!对象属性)
        reportImgTag.lngStartVer = Val(rsData!开始版)
        reportImgTag.strKey = Val(rsData!ID)
        reportImgTag.strImgMarks = ""
        
        If Val(aryImgPro(0)) = 2 Then '报告图
            reportImgTag.lngImgType = ritReport
            
            If blnIsAbort = False Then
                result = ReadReportImage(lngAdviceId, dcmReportImg.Images, reportImgTag, blnMoveState)
            Else
                '加载替换图像
                Set objDcmImg = dcmReportImg.Images.AddNew
                
                dcmReportImg.Images(dcmReportImg.Images.Count).tag = reportImgTag
                
                Call DrawBorder(objDcmImg, 0)
                Call DrawErrorText(objDcmImg, "已被终止")
                
            End If
            
            Call CalcImgView
            
            If result = frAbort Then
                '如果下载异常，且选择终止下载，则退出图像加载处理
                blnIsAbort = True
            End If
        End If
        
        Call rsData.MoveNext
    Wend
End Sub


Private Sub CalcImgView()
    Dim iCols As Integer, iRows As Integer
    
    If dcmReportImg.Images.Count = 1 Then Exit Sub
    
On Error Resume Next
      
    '调整图像显示布局
    ResizeRegion dcmReportImg.Images.Count, dcmReportImg.Width, dcmReportImg.Height, iRows, iCols

    dcmReportImg.MultiColumns = iCols
    dcmReportImg.MultiRows = iRows
    
    If dcmReportImg.Images.Count > 0 Then
        dcmReportImg.CurrentIndex = 1
    Else
        dcmReportImg.CurrentIndex = 0
    End If
End Sub


Private Function ReadReportImage(ByVal lngAdviceId As Long, _
    objImages As DicomImages, reportImgTag As TReportImgTag, ByVal blnMoveState As Boolean) As ftpResult
'读取报告图
    Dim strFile As String
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim objPicMarks As clsPicMarks
    Dim dblMarkZoom As Double
    Dim strError As String
    Dim objDcmImg As DicomImage
    Dim strFileName As String
    Dim blnImgReadState As Boolean
    Dim strReportImgPath As String
    Dim lngImgAdviceId As Long
    
    
    ReadReportImage = frNormal
    blnImgReadState = True
    
    strFileName = GetReportImagePro(reportImgTag.strPros, "PicName")
    lngImgAdviceId = Val(GetReportImagePro(reportImgTag.strPros, "ADVICEID"))
    
    strReportImgPath = GetReportImgPath(lngAdviceId, blnMoveState)
    
    If DirExists(strReportImgPath) = False Then Call MkLocalDir(strReportImgPath)
    
    If Len(strFileName) > 0 Then
        
        strFile = FormatFilePath(strReportImgPath & "\" & strFileName)
        
        '从ftp下载图像
        If FileExists(strFile) = False Then
            If lngImgAdviceId = 0 Then lngImgAdviceId = lngAdviceId
            
            ReadReportImage = DownLoadFtpFile(lngImgAdviceId, strFileName, strFile, blnMoveState)
            If ReadReportImage <> frNormal Then
                blnImgReadState = False
            End If
        End If
    Else
        strFile = FormatFilePath(strReportImgPath & "\报告图_" & reportImgTag.strKey & ".JPG")
        
        '从数据库读取图像
        If FileExists(strFile) = False Then
            Call Sys.ReadLob(glngSys, 6, reportImgTag.strKey, strFile)
        End If
    End If
    
    If FileExists(strFile) = False Then
        If Len(strError) <= 0 Then strError = "未找到对应的报告图像文件 [" & strFile & "]"
        blnImgReadState = False
    End If
    
    If blnImgReadState Then
        '图像读取成功的处理
        Set objDcmImg = ReadDicomFile(strFile, strError)
        
        If Not objDcmImg Is Nothing Then
            reportImgTag.strImgFile = strFileName
            
            objDcmImg.tag = reportImgTag
            
            Call objImages.Add(objDcmImg)
            Call DrawBorder(objDcmImg, 0)
        Else
            blnImgReadState = False
        End If
    End If
    
    If blnImgReadState = False Then
        '加载失败的图像
        
        Set objDcmImg = objImages.AddNew
        
        objImages(objImages.Count).tag = reportImgTag
        
        Call DrawBorder(objDcmImg, 0)
        Call DrawErrorText(objDcmImg, strError)
        
        If ReadReportImage = frNormal Then Call MsgboxH(hwnd, "图像读取失败。" & vbCrLf & strError, vbOKOnly, "提示")
    End If
End Function

Public Function GetLayoutStr() As String
'返回格式字符串[Key=TESTNAME@picturebox1.width:20;picturebox1.height:30;]
    GetLayoutStr = "[KEY=HISTORY@" & _
                                        GetProFmt("VSFSTUDY.HEIGHT", vsfStudy.Height) & _
                                        GetProFmt("PICCONTEXT.HEIGHT", picContext.Height) & _
                                 "]"
End Function

Public Sub SetLayout(ByVal strLayout As String)
    Dim strPros As String
    Dim strPro As String
    
    If Len(strLayout) <= 0 Then Exit Sub
    
    strPros = GetPros(strLayout, "HISTORY")
    
    strPro = GetProValue(strPros, "VSFSTUDY.HEIGHT")
    If Val(strPro) > 0 Then vsfStudy.Height = Val(strPro)
    
    strPro = GetProValue(strPro, "PICCONTEXT.HEIGHT")
    If Val(strPro) > 0 Then picContext.Height = Val(strPro)
    

End Sub


Private Function DownLoadFtpFile(ByVal lngAdviceId As Long, _
    ByVal strFtpFile As String, ByVal strLocalFile As String, ByVal blnMoveState As Boolean) As ftpResult
'下载ftp文件
    DownLoadFtpFile = frNormal
    If Len(mftpConTag.Ip) <= 0 Or Val(mftpConTag.tag) <> lngAdviceId Then
        mftpConTag = GetReportDevice(lngAdviceId, blnMoveState)
        mftpConTag.tag = lngAdviceId
        
        If Len(mftpConTag.Ip) <= 0 Then
            DownLoadFtpFile = frAbort
            Exit Function
        End If
    End If
    
    DownLoadFtpFile = FtpDownload(mftpConTag, strFtpFile, strLocalFile)
End Function


Private Function GetReportDevice(ByVal lngAdviceId As Long, ByVal blnMoveState As Boolean) As TFtpConTag
    Dim strSQL As String
    Dim rsData As ADODB.Recordset
    
On Error GoTo errhandle
    strSQL = "select NVl(相关ID, ID) as ID from 病人医嘱记录 where ID=[1]"
    
    If blnMoveState Then strSQL = Replace(strSQL, "病人医嘱记录", "H病人医嘱记录")
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询主医嘱ID", lngAdviceId)
    
    If rsData.RecordCount <= 0 Then
        Call MsgboxH(hwnd, "医嘱数据校验失败，未找到报告关联医嘱信息。", vbOKOnly, "提示")
        Exit Function
    End If
    
    strSQL = " Select Decode(A.接收日期,Null,'',to_Char(A.接收日期,'YYYYMMDD')||'/') ||A.检查UID||'/' As URL," & _
            " B.设备号 as 设备号1, B.设备名 As 设备名1, B.FTP用户名 As User1,B.FTP密码 As Pwd1, B.IP地址 As Host1, " & _
                    " decode(B.Ftp目录, null, '/', '/'||B.Ftp目录||'/') As Root1,B.共享目录 as 共享目录1,B.共享目录用户名 as 共享目录用户名1,B.共享目录密码 as 共享目录密码1 " & _
            " From  影像检查记录 A,影像设备目录 B " & _
            " Where A.医嘱ID=[1] And nvl(A.位置一, A.位置二)=B.设备号(+)  "
            
    If blnMoveState Then strSQL = Replace(strSQL, "影像检查记录", "H影像检查记录")
    Set rsData = zlDatabase.OpenSQLRecord(strSQL, "查询报告图像存储", Val(rsData!ID))
            
    If rsData.RecordCount <= 0 Then
        Call MsgboxH(hwnd, "未找到报告图对应的存储设备，请检查数据是否正确。", vbOKOnly, "提示")
        Exit Function
    End If
    
    If nvl(rsData!Host1) <> "" Then
        GetReportDevice = FtpTagInstance(rsData!Host1, rsData!User1, rsData!Pwd1, rsData!Root1 & rsData!Url)
    End If
Exit Function
errhandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub SelectedAll()
'全选
    Dim i As Long
    
    For i = 1 To dcmReportImg.Images.Count
        Call DrawBorder(dcmReportImg.Images(i), ColorConstants.vbRed, True)
    Next i
End Sub

Private Sub cmdWrite_Click()
On Error GoTo errhandle
    Call WriteReport
Exit Sub
errhandle:
    Call MsgboxH(hwnd, err.Description, vbOKOnly, "提示")
End Sub

Public Sub WriteReport()
'写入报告
    If vsfStudy.Rows <= 0 Then Exit Sub
    
    RaiseEvent OnSend
End Sub

Private Sub dcmReportImg_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo errhandle
    Dim lngSelectIndex As Long
    Dim i As Long
    
    If dcmReportImg.Images.Count <= 0 Then Exit Sub
    
    Select Case KeyCode
        Case 37     '左光标键盘
            lngSelectIndex = mlngSelImgIndex - 1
            If lngSelectIndex <= 0 Then Exit Sub
        Case 38    '上光标键
            lngSelectIndex = mlngSelImgIndex - dcmReportImg.MultiColumns
            If lngSelectIndex <= 0 Then Exit Sub
        Case 39      '右光标键
            lngSelectIndex = mlngSelImgIndex + 1
            If lngSelectIndex > dcmReportImg.Images.Count Then Exit Sub
        Case 40      '下光标键
            lngSelectIndex = mlngSelImgIndex + dcmReportImg.MultiColumns
            If lngSelectIndex > dcmReportImg.Images.Count Then Exit Sub
        Case 65
            If Shift = 2 Then
                Call SelectedAll  '按下全选
                lngSelectIndex = 0
                Exit Sub
            End If
            
        Case Else
            Exit Sub
    End Select
    
    For i = 1 To dcmReportImg.Images.Count
        Call DrawBorder(dcmReportImg.Images(i), 0)
    Next
        
    If lngSelectIndex > 0 Then
        Call DrawBorder(dcmReportImg.Images(lngSelectIndex), ColorConstants.vbRed, True)
    End If
    
    mlngSelImgIndex = lngSelectIndex
     
Exit Sub
errhandle:
    Call MsgboxH(hwnd, err.Description, vbOKOnly, "提示")
End Sub

Private Sub dcmReportImg_MouseUp(Button As Integer, Shift As Integer, X As Long, Y As Long)
On Error GoTo errhandle
    Dim i As Long
    
    If Button = 2 Then
        '鼠标右键
    Else
        mlngSelImgIndex = dcmReportImg.ImageIndex(X, Y)
        
        If mlngSelImgIndex <= 0 Or mlngSelImgIndex > dcmReportImg.Images.Count Then Exit Sub
        
        If Shift <> 2 Then
            For i = 1 To dcmReportImg.Images.Count
                Call DrawBorder(dcmReportImg.Images(i), 0)
            Next
        End If
            
        Call DrawBorder(dcmReportImg.Images(mlngSelImgIndex), ColorConstants.vbRed, True)
    End If
    
    RaiseEvent OnMouseUp(Button, Shift, CSng(X), CSng(Y))
Exit Sub
errhandle:
    MsgboxH hwnd, err.Description, vbOKOnly, "提示"
End Sub

Private Sub picContext_Resize()
On Error GoTo errhandle
    chkReportType.Left = 0
    chkReportType.Top = 0
    
    cmdWrite.Top = 0
    cmdWrite.Left = chkReportType.Left + chkReportType.Width ' picContext.ScaleWidth - cmdWrite.Width
    
    chkLinkView.Left = picContext.Width - chkLinkView.Width
    chkLinkView.Top = 45
    
    rtxtReport.Left = 0
    rtxtReport.Top = chkReportType.Height
    rtxtReport.Width = picContext.ScaleWidth
    rtxtReport.Height = picContext.ScaleHeight - chkReportType.Height
    
    dcmReportImg.Left = 0
    dcmReportImg.Top = chkReportType.Height
    dcmReportImg.Width = picContext.ScaleWidth
    dcmReportImg.Height = picContext.ScaleHeight - chkReportType.Height
Exit Sub
errhandle:

End Sub

Private Sub rtxtReport_SelChange()
On Error GoTo errhandle
    cmdWrite.Enabled = IIf(rtxtReport.SelLength > 0, True, False)
Exit Sub
errhandle:

End Sub

Private Sub UserControl_Initialize()
    mstrDateRange = "一年"
    mblnAllowWrite = True
End Sub

Private Sub UserControl_Resize()
On Error GoTo errhandle
    picBack.Move 0, 0, ScaleWidth, ScaleHeight
    
    Call ucSplitter1.RePaint(False)
Exit Sub
errhandle:

End Sub
 
Private Sub UserControl_Terminate()
    Call Destory
    
    Set mobjOwner = Nothing
End Sub

'Private Sub UserControl_Show()
'On Error GoTo errHandle
'    If UserControl.Ambient.UserMode Then
'        Call Init(mlngCurModule, mlngCurDeptID)
'    End If
'Exit Sub
'errHandle:
'
'End Sub



Private Sub vsfStudy_AfterMoveColumn(ByVal Col As Long, Position As Long)
On Error GoTo errhandle
    mTPListCfg.strColPros = GetListHeadString
    
    SaveSetting "ZLSOFT", "私有模块\" & UserInfo.用户名 & "\" & App.ProductName & "\ReportHistory", "列表配置", mTPListCfg.strSortPro & ";" & mTPListCfg.strColPros
Exit Sub
errhandle:

End Sub

Private Sub vsfStudy_AfterSort(ByVal Col As Long, Order As Integer)
    Dim strName As String
    Dim i As Integer
    
On Error GoTo errhandle
    For i = 1 To vsfStudy.Rows - 1
        vsfStudy.TextMatrix(i, vsfStudy.ColIndex("序号")) = i
    Next
    
    strName = vsfStudy.TextMatrix(0, Col)
    mTPListCfg.strSortPro = strName & "," & Order
    
    mTPListCfg.strColPros = GetListHeadString
    
    SaveSetting "ZLSOFT", "私有模块\" & UserInfo.用户名 & "\" & App.ProductName & "\ReportHistory", "列表配置", mTPListCfg.strSortPro & ";" & mTPListCfg.strColPros
Exit Sub
errhandle:

End Sub

Private Sub vsfStudy_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
On Error GoTo errhandle
    mTPListCfg.strColPros = GetListHeadString
    
    SaveSetting "ZLSOFT", "私有模块\" & UserInfo.用户名 & "\" & App.ProductName & "\ReportHistory", "列表配置", mTPListCfg.strSortPro & ";" & mTPListCfg.strColPros
Exit Sub
errhandle:
End Sub


Private Sub vsfStudy_Click()
On Error GoTo errhandle
     
    RaiseEvent OnClick
    
    Exit Sub
errhandle:
    MsgboxH hwnd, err.Description, vbExclamation, "提示"
    err.Clear
End Sub

Private Sub vsfStudy_DblClick()
On Error GoTo errhandle

    If chkLinkView.Visible Then Call DoLinkView(True)
    
    '可通过双击，弹出历次报告明细显示窗口（包含历次检查记录显示，报告内容显示，医嘱显示，病历显示等）
    RaiseEvent OnDbClick
    
    Exit Sub
errhandle:
    MsgboxH hwnd, err.Description, vbExclamation, "提示"
    err.Clear
End Sub


Private Sub vsfStudy_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo errH
     
    RaiseEvent OnMouseUp(Button, Shift, X, Y)
    
errH:
End Sub

Public Sub ClearData()
'清空数据
    vsfStudy.Rows = 1
    
    mlngAdviceId = 0
End Sub


Public Function IsImageEnable(ByVal lngAdvice As Long) As Boolean
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    If lngAdvice <= 0 Then
        IsImageEnable = False
        Exit Function
    End If
    
    strSQL = "select 检查UID from 影像检查记录 where  医嘱id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询检查UID", lngAdvice)
    
    If rsTemp.RecordCount <= 0 Then Exit Function
    
    IsImageEnable = Len(nvl(rsTemp!检查UID)) > 0
End Function

Public Function IsReportEnable(ByVal lngAdvice As Long) As Boolean
    Dim strSQL As String
    Dim rsTemp As Recordset
    
    If lngAdvice <= 0 Then
        IsReportEnable = False
        Exit Function
    End If
    
    strSQL = "select count(1) 计数 from 病人医嘱报告 where  医嘱id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询报告", lngAdvice)
    
    If rsTemp.RecordCount <= 0 Then Exit Function
    
    IsReportEnable = Val(nvl(rsTemp!计数)) > 0
End Function

Public Sub Destory()
On Error GoTo errhandle

    ucSplitter1.Destory
    
Exit Sub
errhandle:
    Debug.Print "ucReportHistory_Destory Err:" & err.Description
End Sub

Private Sub vsfStudy_SelChange()
On Error GoTo errhandle
    Dim intCol As Integer
    Dim lngAdviceId As Long
    Dim blnMoved As Boolean
    
    If vsfStudy.Rows <= 1 Then Exit Sub
    If vsfStudy.Row <= 0 Then Exit Sub
    
    intCol = vsfStudy.ColIndex("医嘱ID")
    If intCol = -1 Then Exit Sub
    
    lngAdviceId = Val(vsfStudy.TextMatrix(vsfStudy.Row, intCol))
    mftpConTag.Ip = ""
    
    blnMoved = IIf(Val(vsfStudy.TextMatrix(vsfStudy.Row, vsfStudy.ColIndex("转储状态"))) = 0, False, True)
    '显示报告内容...
    Call LoadReport(lngAdviceId, _
        vsfStudy.TextMatrix(vsfStudy.Row, vsfStudy.ColIndex("加载类别")), _
        vsfStudy.TextMatrix(vsfStudy.Row, vsfStudy.ColIndex("关键ID")), _
        blnMoved)
        
    rtxtReport.tag = lngAdviceId
    
    rtxtReport.Visible = True
    dcmReportImg.Visible = False
    dcmReportImg.Images.Clear
    dcmReportImg.tag = 0
    
    chkReportType.value = Unchecked
    
    '联动查看
    If chkLinkView.Visible And chkLinkView.value <> 0 Then Call DoLinkView(False)
    
Exit Sub
errhandle:
    MsgboxH GetRootHwnd, err.Description, vbOKOnly, "提示"
End Sub


Private Sub LoadProcessReport(ByVal lngKey As Long, ByVal strFormatHead As String, ByVal strFontSize As String, _
    ByVal blnMovedState As Boolean)
'载入病理过程报告内容
'lngReportId 过程报告ID
'strFormatHead 格式头
'strFontSize 字体大小

    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strFormatContext  As String
    Dim strtext As String
    Dim strTitle As String
    Dim strAttachInfo As String
    
    On Error GoTo errH
    strFormatContext = strFormatHead
    
    
    '显示附加内容
    strAttachInfo = LoadAttachInfo(mlngAdviceId, strFontSize, blnMovedState)
    strFormatContext = strFormatContext & strAttachInfo
    
    '查询过程报告内容
    strSQL = "select 检查结果,检查意见 from 病理过程报告 where id=[1]"

    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "历史病理过程报告查询", lngKey)
                
    If rsTemp.RecordCount > 0 Then
    
        If Trim(strAttachInfo) <> "" Then
            strFormatContext = strFormatContext & "\b\cf0\fs" & strFontSize & "==**************************==" & "\par"
        End If
        
        strTitle = "检查结果" & "："
        strtext = nvl(rsTemp!检查结果) & vbCrLf
        strFormatContext = strFormatContext & "\b\cf2\fs24 " & strTitle & " \par\b0\cf0\fs" & strFontSize & " " & Replace(strtext, vbCrLf, " \par\cf0\fs" & strFontSize & " ") & "\par"
        
        strTitle = "检查意见" & "："
        strtext = nvl(rsTemp!检查意见) & vbCrLf
        strFormatContext = strFormatContext & "\b\cf2\fs24 " & strTitle & " \par\b0\cf0\fs" & strFontSize & " " & Replace(strtext, vbCrLf, " \par\cf0\fs" & strFontSize & " ") & "\par"
            
        strFormatContext = strFormatContext & "}"
        rtxtReport.SelRTF = strFormatContext
        rtxtReport.SelStart = 0
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
End Sub

Private Sub LoadDescription(ByVal lngKey As Long, ByVal strFormatHead As String, ByVal strFontSize As String, _
    ByVal blnMovedState As Boolean)
'载入巨检描述内容  mstrPatholMaterialInfo 标本名称,取材位置,形状,蜡块数,制片数,主取医师,取材时间,性质,颜色,标本量
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strFormatContext  As String
    Dim strtext As String
    Dim strTitle As String
    Dim str巨检描述 As String
    Dim blnIsCell As Boolean '材块是否是细胞类型 细胞类型是2
    Dim strAttachInfo As String
    
    On Error GoTo errH
    
    strFormatContext = strFormatHead
    
    '显示附加内容
    strAttachInfo = LoadAttachInfo(mlngAdviceId, strFontSize, blnMovedState)
    strFormatContext = strFormatContext & strAttachInfo
    
    '查询巨检描述
    strSQL = "select  a.巨检描述,a.检查类型,b.序号,b.标本名称, b.形状,b.取材位置,b.蜡块数,b.主取医师,b.性质,b.颜色,b.标本量,b.取材时间, b.标本名称, c.制片数 " & _
                      "from 病理检查信息 a ,病理取材信息 b ,病理制片信息 c " & _
                      "where b.材块id=c.材块id and a.病理医嘱id=c.病理医嘱id and a.病理医嘱id=b.病理医嘱id and a.病理医嘱id=[1] order by b.序号 "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "历史巨检描述查询", lngKey)
        
    If rsTemp.RecordCount > 0 Then
        str巨检描述 = nvl(rsTemp!巨检描述)
        blnIsCell = (Val(nvl(rsTemp!检查类型)) = 2)   '细胞类型是2
        
        If Trim(strAttachInfo) <> "" Then
            strFormatContext = strFormatContext & "\b\cf0\fs" & strFontSize & "==**************************==" & "\par"
        End If
    End If
    
    If UBound(Split(mstrPatholMaterialInfo, ",")) <> 9 Then mstrPatholMaterialInfo = "1,1,1,1,1,1,1,1,1,1"
                
    While Not rsTemp.EOF
    
        strTitle = "蜡块" & nvl(rsTemp!序号) & "："
        strtext = ""
        
        If Split(mstrPatholMaterialInfo, ",")(0) = 1 And Trim(nvl(rsTemp!标本名称)) <> "" Then strtext = "标本名称：" & nvl(rsTemp!标本名称)
        If Split(mstrPatholMaterialInfo, ",")(1) = 1 And Trim(nvl(rsTemp!取材位置)) <> "" Then strtext = IIf(strtext <> "", strtext & "，" & "取材位置：" & nvl(rsTemp!取材位置), "取材位置：" & nvl(rsTemp!取材位置))
        If Split(mstrPatholMaterialInfo, ",")(2) = 1 And Trim(nvl(rsTemp!形状)) <> "" Then strtext = IIf(strtext <> "", strtext & "，" & "形状：" & nvl(rsTemp!形状), "形状：" & nvl(rsTemp!形状))
        
        If blnIsCell Then
            If Split(mstrPatholMaterialInfo, ",")(7) = 1 And Trim(nvl(rsTemp!性质)) <> "" Then strtext = IIf(strtext <> "", strtext & "，" & "性质：" & nvl(rsTemp!性质), "性质：" & nvl(rsTemp!性质))
            If Split(mstrPatholMaterialInfo, ",")(8) = 1 And Trim(nvl(rsTemp!颜色)) <> "" Then strtext = IIf(strtext <> "", strtext & "，" & "颜色：" & nvl(rsTemp!颜色), "颜色：" & nvl(rsTemp!颜色))
            If Split(mstrPatholMaterialInfo, ",")(9) = 1 And Trim(nvl(rsTemp!标本量)) <> "" Then strtext = IIf(strtext <> "", strtext & "，" & "标本量：" & nvl(rsTemp!标本量), "标本量：" & nvl(rsTemp!标本量))
            If Split(mstrPatholMaterialInfo, ",")(3) = 1 And Trim(nvl(rsTemp!蜡块数)) <> "" Then strtext = IIf(strtext <> "", strtext & "，" & "细胞块数：" & nvl(rsTemp!蜡块数), "细胞块数：" & nvl(rsTemp!蜡块数))
        Else
            If Split(mstrPatholMaterialInfo, ",")(3) = 1 And Trim(nvl(rsTemp!蜡块数)) <> "" Then strtext = IIf(strtext <> "", strtext & "，" & "材块数：" & nvl(rsTemp!蜡块数), "材块数：" & nvl(rsTemp!蜡块数))
        End If
        
        If Split(mstrPatholMaterialInfo, ",")(4) = 1 And Trim(nvl(rsTemp!制片数)) <> "" Then strtext = IIf(strtext <> "", strtext & "，" & "制片数：" & nvl(rsTemp!制片数), "制片数：" & nvl(rsTemp!制片数))
        If Split(mstrPatholMaterialInfo, ",")(5) = 1 And Trim(nvl(rsTemp!主取医师)) <> "" Then strtext = IIf(strtext <> "", strtext & "，" & "主取医师：" & nvl(rsTemp!主取医师), "主取医师：" & nvl(rsTemp!主取医师))
        If Split(mstrPatholMaterialInfo, ",")(6) = 1 And Trim(nvl(rsTemp!取材时间)) <> "" Then strtext = IIf(strtext <> "", strtext & "，" & "取材时间：" & nvl(rsTemp!取材时间), "取材时间：" & nvl(rsTemp!取材时间))
        
        If strtext <> "" Then
            strtext = strtext & vbCrLf
            strFormatContext = strFormatContext & "\b\cf2\fs24 " & strTitle & " \par\b0\cf0\fs" & strFontSize & " " & Replace(strtext, vbCrLf, " \par\cf0\fs" & strFontSize & " ") & "\par"
        End If
        
        rsTemp.MoveNext
    Wend
    
    If Trim(str巨检描述) <> "" Then strFormatContext = strFormatContext & "\b\cf2\fs24 " & "巨检描述:" & " \par\b0\cf0\fs" & strFontSize & " " & Replace(str巨检描述, vbCrLf, " \par\cf0\fs" & strFontSize & " ") & "\par"
        
    strFormatContext = strFormatContext & "}"
    rtxtReport.SelRTF = strFormatContext
    rtxtReport.SelStart = 0
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
End Sub

Private Function LoadAttachInfo(ByVal lngAdviceId As Long, ByVal strFontSize As String, ByVal blnMovedState As Boolean) As String
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strResult As String
    
    strResult = ""
    
    '显示附加内容
    strSQL = "Select 项目,内容 From 病人医嘱附件 Where 医嘱ID=[1] Order By 排列"
    If blnMovedState Then
        strSQL = Replace(strSQL, "病人医嘱附件", "H病人医嘱附件")
    End If

    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "提取病人附件", mlngAdviceId)
    Do Until rsTemp.EOF
        If nvl(rsTemp!项目) <> "" And nvl(rsTemp!内容) <> "" Then
            strResult = strResult & "\b\cf0\fs" & strFontSize & " " & rsTemp!项目 & ":" & " \b0\cf0\fs" & strFontSize & " " & Replace(nvl(rsTemp!内容), vbCrLf, " \par\cf0\fs24 ") & "\par"
        End If
        rsTemp.MoveNext
    Loop

    strSQL = "select 信息名,信息值 from 病人信息从表 where 病人ID=[1] and 就诊id=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "提取外院病人信息", mlngPatientId, mlngAdviceId)
    Do Until rsTemp.EOF
        If nvl(rsTemp!信息名) <> "" And nvl(rsTemp!信息值) <> "" Then
            strResult = strResult & "\b\cf0\fs" & strFontSize & " " & rsTemp!信息名 & ":" & " \b0\cf0\fs" & strFontSize & " " & Replace(nvl(rsTemp!信息值), vbCrLf, " \par\cf0\fs24 ") & "\par"
        End If
        rsTemp.MoveNext
    Loop
    
    LoadAttachInfo = strResult
End Function

Private Sub LoadReportContent(ByVal lngKey As Long, _
    ByVal strFormatHead As String, _
    ByVal strFontSize As String, _
    ByVal blnMovedState As Boolean)
'载入报告内容
'lngReportId 报告ID
'strFormatHead 格式头
'strFontSize 字体大小

    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim blnShow As Boolean
    Dim strFormatContext  As String
    Dim strtext As String
    Dim strTitle As String
    Dim strAttachInfo As String
    
    On Error GoTo errH
    
    strFormatContext = strFormatHead
    
    '显示附加内容
    strAttachInfo = LoadAttachInfo(mlngAdviceId, strFontSize, blnMovedState)
    strFormatContext = strFormatContext & strAttachInfo

    '读取报告的内容
    strSQL = "Select a.内容文本 As 标题, b.对象属性, b.内容文本 As 正文,b.开始版 as 版本 From 电子病历内容 a,电子病历内容 b " & _
             " Where a.文件id = [1] And a.对象类型 = 3 And a.Id = b.父ID And b.对象类型 = 2 and b.终止版=0  "
    
    If blnMovedState Then
        strSQL = Replace(strSQL, "电子病历内容", "H电子病历内容")
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "历史报告内容查询", lngKey)
    
    If rsTemp.RecordCount > 0 And Trim(strAttachInfo) <> "" Then
        strFormatContext = strFormatContext & "\par\b\cf0\fs" & strFontSize & "==**************************==" & "\par"
    End If
    
    If rsTemp.RecordCount <= 0 Then
        If Len(strAttachInfo) <= 0 Then
            strFormatContext = strFormatContext & "\b\cf1\fs" & strFontSize & "◆暂无报告..." & "\par"
        Else
            strFormatContext = strFormatContext & "\par\b\cf1\fs" & strFontSize & "◆暂无报告..." & "\par"
        End If
    End If
                
    While Not rsTemp.EOF
        blnShow = False
        Select Case rsTemp!标题
            Case "检查所见"
                strTitle = mstrDescriptionName
                strtext = nvl(rsTemp!正文) & vbCrLf
                blnShow = True
            Case "诊断意见"
                strTitle = mstrOpinionName
                strtext = nvl(rsTemp!正文) & vbCrLf
                blnShow = True
            Case "建议"
                strTitle = mstrAdviseName
                strtext = nvl(rsTemp!正文) & vbCrLf
                blnShow = True
        End Select
        
        If blnShow = True Then strFormatContext = strFormatContext & "\b\cf2\fs24 " & strTitle & " \par\b0\cf0\fs" & strFontSize & " " & Replace(strtext, vbCrLf, " \par\cf0\fs" & strFontSize & " ") & "\par"
        rsTemp.MoveNext
    Wend
    
    strFormatContext = strFormatContext & "}"
    rtxtReport.SelRTF = strFormatContext
    rtxtReport.SelStart = 0
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
End Sub


Private Sub LoadReport(ByVal lngAdviceId As Long, _
    ByVal strLoadType As String, ByVal lngKeyId As Long, _
    ByVal blnMoveState As Boolean)
'本过程中lvHistoryList.ListItems关键字分为 process：过程报告 ；describe：巨检描述；其他：检查所见意见建议等内容（原来使用的K）
On Error GoTo err
    Dim strSQL As String
    Dim strtext As String
    Dim strFormatContext As String
    Dim strSize As String
    Dim rsTemp As ADODB.Recordset
    
    rtxtReport.Text = ""
    
    strSize = FontSize
    strSize = 2 * Round(Val(strSize))
    

    strFormatContext = "{\rtf1\ansi\ansicpg936\deff0\deflang1033\deflangfe2052{\fonttbl{\f0\fnil\fcharset134 \'cb\'ce\'cc\'e5;}}" & _
                       "{\colortbl ;\red255\green104\blue104;\red19\green164\blue251;}" & _
                       "{\*\generator Msftedit 5.41.21.2509;}\viewkind4\uc1\sl276\slmult1\lang2052\b\f0\fs24 "
         
    If InStr(strLoadType, M_STR_LISTVIEWKET_PROCESS) > 0 Then
        Call LoadProcessReport(lngKeyId, strFormatContext, strSize, blnMoveState)
    ElseIf InStr(strLoadType, M_STR_LISTVIEWKEY_DESCRIBE) > 0 Then
        Call LoadDescription(lngKeyId, strFormatContext, strSize, blnMoveState)
    Else
        '根据医嘱id，查询对应报告ID
        strSQL = "select 病历ID from 病人医嘱报告 where 医嘱ID=[1] "
        If blnMoveState Then
            strSQL = Replace(strSize, "病人医嘱报告", "H病人医嘱报告")
        End If
        
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "历史医嘱报告ID查询", lngAdviceId)
        
        If rsTemp.RecordCount > 0 Then
            Call LoadReportContent(Val(nvl(rsTemp!病历Id)), strFormatContext, strSize, blnMoveState)
        Else
            Call LoadReportContent(0, strFormatContext, strSize, blnMoveState)
        End If
        
    End If

    Exit Sub
err:
    If ErrCenter = 1 Then
        Resume
    Else
        Call MsgboxH(GetRootHwnd, err.Description, vbOKOnly, "提示")
    End If
End Sub


Private Function GetListHeadString() As String
'得到列名参数: 名称,宽度,是否显示  例如  "类别,1000,1|执行过程,2000,0|"
On Error GoTo errH
    Dim i As Integer
    Dim strTemp As String
    Dim strName As String
    Dim lngWidth As Long
    Dim blnIsHide As Boolean
    
    For i = 0 To vsfStudy.Cols - 1
        
        strName = vsfStudy.TextMatrix(0, i)
        
        lngWidth = vsfStudy.ColWidth(i)
        blnIsHide = vsfStudy.ColHidden(i)
        
        If Len(strTemp) > 0 Then
            strTemp = strTemp & "|"
        End If
        
        strTemp = strTemp & strName & "," & lngWidth & "," & blnIsHide
    Next

    GetListHeadString = strTemp
    
    Exit Function
errH:
    err.Raise -1, "历史检查", "[获取列头配置]" & vbCrLf & err.Description
    Resume
End Function

Private Sub DoLoadListSort(ByVal strcfg As String)
'恢复排序
On Error GoTo errH
    Dim strName As String
    Dim intWay As Integer
    Dim intPos As Integer
    Dim intCol As Integer
    Dim i As Integer
    
    intPos = InStr(strcfg, ",")
    If intPos = 0 Then Exit Sub
    
    strName = Split(strcfg, ",")(0)
    intWay = Val(Split(strcfg, ",")(1))
    
    With vsfStudy
        For i = 1 To .Cols - 1
            If strName = .TextMatrix(0, i) Then
                intCol = i
                Exit For
            End If
        Next
         
        .Col = intCol
        .Sort = intWay
        
        For i = 1 To .Rows - 1
            .TextMatrix(i, vsfStudy.ColIndex("序号")) = i
        Next
    End With
    
    Exit Sub
errH:
    err.Raise -1, "列表个性化设置", "[DoLoadListSort]" & vbCrLf & err.Description
    Resume
End Sub





