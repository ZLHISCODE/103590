VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.Unicode.9600.ocx"
Begin VB.Form frmSentenceList 
   BorderStyle     =   0  'None
   Caption         =   "词句示范列表"
   ClientHeight    =   7590
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   ScaleHeight     =   7590
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin XtremeReportControl.ReportControl rptList 
      Height          =   2085
      Left            =   165
      TabIndex        =   1
      Top             =   2385
      Width           =   3000
      _Version        =   589884
      _ExtentX        =   5292
      _ExtentY        =   3678
      _StockProps     =   0
      BorderStyle     =   2
      MultipleSelection=   0   'False
      EditOnClick     =   0   'False
   End
   Begin MSComctlLib.TreeView TreeList 
      Height          =   1470
      Left            =   315
      TabIndex        =   4
      Top             =   705
      Width           =   2310
      _ExtentX        =   4075
      _ExtentY        =   2593
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "imgClass"
      Appearance      =   0
   End
   Begin VB.PictureBox picTerm 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   4050
      Left            =   3330
      ScaleHeight     =   4050
      ScaleWidth      =   2445
      TabIndex        =   2
      Top             =   435
      Visible         =   0   'False
      Width           =   2445
      Begin VSFlex8Ctl.VSFlexGrid vfgTerm 
         Height          =   3690
         Left            =   60
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   120
         Width           =   2340
         _cx             =   4128
         _cy             =   6509
         Appearance      =   2
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   16761024
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483643
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   16777215
         GridColor       =   -2147483643
         GridColorFixed  =   -2147483643
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   2
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   4
         Cols            =   1
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
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
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
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
   Begin MSComctlLib.ImageList imgList 
      Left            =   750
      Top             =   4320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSentenceList.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSentenceList.frx":059A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSentenceList.frx":0B34
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSentenceList.frx":10CE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtbText 
      Height          =   1755
      Left            =   180
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   4695
      Width           =   2955
      _ExtentX        =   5212
      _ExtentY        =   3096
      _Version        =   393217
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmSentenceList.frx":19A8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList imgClass 
      Left            =   2025
      Top             =   300
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
            Picture         =   "frmSentenceList.frx":1A45
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSentenceList.frx":1FDF
            Key             =   "expend"
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgListTmp 
      Height          =   420
      Left            =   3960
      TabIndex        =   5
      Top             =   5775
      Visible         =   0   'False
      Width           =   585
      _cx             =   1032
      _cy             =   741
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
      BackColorFixed  =   15790320
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16772055
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
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   300
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
      AutoResize      =   0   'False
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
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Left            =   855
      Top             =   105
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   195
      Top             =   30
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "frmSentenceList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'######################################################################################################################

'常量定义
'------------------------------------------------------------------------------------------------------------------
Private Enum mCol
    图标 = 0: ID: 分类id: 分类: 编号: 名称: 部门: 人员
End Enum

Private Const con_UnDefine = -999
Private Const conPane_Tree = 400
Private Const conPane_List = 401
Private Const conPane_Term = 403
Private Const conPane_Text = 404


'窗体变量
'------------------------------------------------------------------------------------------------------------------
Private mfrmParent As Form          '当前窗体的上级窗体
Private mstrPrivs As String         '当前使用者权限串
Private mlngWordId As Long          '当前词句id
Private mblnCompend As Boolean      '按提纲列举词句：维护管理程序按分类列举，病历编辑中按提纲列举
Private mlngParentId As Long        '父id，当为单分类时，为分类Id直接匹配的分类id，当为多分类时，为病历文件结构的提纲id
Private mlngClassId As Long         '词句默认分类id,当按分类列举时，和mlngParentId相同，当按提纲列举时，为对应的一个分类id
Private mlngPatient As Long          '病人id，在病人病历编辑时，用来确定条件词句是否满足
Private mlngVisit As Long           '主页id或挂号单ID
Private mlngAdvice As Long          '医嘱ID
Private mstrSecondLimit As String   '进行二次过滤的词句ID串，以豆号分隔

Private mintPower As Integer        '词句管理权范围
'    mintPower=con_UnDefine，未定义;
'    mintPower=-1，不具备词句管理权;
'    mintPower=0，全院，这时显示所有的示范，也可以更改;
'    mintPower=1，科室，这时显示全院通用示范(科室id is null)和所在科室公有或部门内人员私有的示范，但不能更改全院通用示范;
'    mintPower=2，个人，这时显示全院通用示范(科室id is null)和所在科室通用示范(人员id is null)和个人示范，仅个人示范可更改
Public Event RowDblClick(ByVal lngSentenceID As Long)    '双击一行或在行上按回车


'以下为外部公共程序
'######################################################################################################################
Public Function zlRefFromClass(ByVal frmParent As Form, ByVal lngClassId As Long) As Long
    '******************************************************************************************************************
    '功能：根据指定分类，刷新列表，属于维护管理接口
    '参数：指定的词句示范分类id
    '******************************************************************************************************************
    Set mfrmParent = frmParent
    mblnCompend = False
    TreeList.Visible = mblnCompend
    picTerm.Visible = Not mblnCompend
    If Not mblnCompend Then
        dkpMan.FindPane(conPane_Tree).Close
        dkpMan.FindPane(conPane_Term).Closed = False
    End If
    If mlngParentId = lngClassId Then zlRefFromClass = Me.rptList.Rows.Count: Exit Function
    mlngParentId = lngClassId
    mlngClassId = lngClassId
    
    rptList.Columns(mCol.部门).Visible = True
    rptList.Columns(mCol.人员).Visible = True
    
    zlRefFromClass = zlSubRefList(mlngWordId)
End Function

Public Function zlRefFromCompend(ByVal frmParent As Form, _
                                ByVal lngCompendID As Long, _
                                Optional lngPatient As Long = 0, _
                                Optional lngVisit As Long = 0, _
                                Optional lngAdvice As Long = 0, _
                                Optional blnForce As Boolean, _
                                Optional strSecondLimit As String) As Long
    '******************************************************************************************************************
    '功能： 根据指定提纲，刷新列表，属于病历编辑接口
    '参数： 指定的文件定义提纲id
    '       lngPatient，病人id
    '       lngVisit，病人就诊ID。门诊病人为挂号ID，住院病人为主页id
    '       lngAdvice，医嘱ID
    '       lngCompendID，提纲id
    '******************************************************************************************************************
    
    Dim rsTemp As New ADODB.Recordset
    Dim panThis As Pane
    
    If blnForce = False And mlngParentId = lngCompendID And _
        (mstrSecondLimit = strSecondLimit) Then zlRefFromCompend = Me.rptList.Rows.Count: Exit Function
    Set mfrmParent = frmParent
    mblnCompend = True
    TreeList.Visible = mblnCompend
    picTerm.Visible = Not mblnCompend
    mlngParentId = lngCompendID
    mlngPatient = lngPatient
    mlngVisit = lngVisit
    mlngAdvice = lngAdvice
    mstrSecondLimit = strSecondLimit
    
    rptList.Columns(mCol.部门).Visible = False
    rptList.Columns(mCol.人员).Visible = False
    Set panThis = dkpMan.FindPane(conPane_Term)
    panThis.Close
    Set panThis = dkpMan.FindPane(conPane_Tree)
    panThis.Closed = False
    
    gstrSQL = "Select 词句分类id From 病历提纲词句 Where 提纲id = [1]"
    Err = 0: On Error GoTo errHand
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngParentId)
    mlngClassId = 0
    If rsTemp.RecordCount > 0 Then mlngClassId = rsTemp.Fields(0).Value
    
    Call zlSubRefClass
    
    zlRefFromCompend = zlSubRefList(mlngWordId, 0)
    
    Exit Function
    
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    
    Call SaveErrLog
    zlRefFromCompend = 0
End Function

Public Sub zlAddFromEditor()
    '******************************************************************************************************************
    '功能：执行指定命令条控件，属于病历编辑接口
    '参数：当前的编辑器窗体
    '******************************************************************************************************************
Dim lngRetuId As Long
Dim cbrControl As CommandBarControl
    
    If mlngClassId = 0 Then
        MsgBox "当前提纲没有设置词句示范分类对应，请联系管理员初始化基础数据！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If mintPower < 0 Then
        MsgBox "你不具备词句示范管理的权限！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    Set cbrControl = Me.cbsThis.ActiveMenuBar.FindControl(xtpControlButton, conMenu_Edit_NewItem, True, True)
    If cbrControl Is Nothing Then Exit Sub
    If cbrControl.Visible = False Or cbrControl.Enabled = False Then Exit Sub
    
    lngRetuId = frmSentenceEdit.ShowMe(mfrmParent, True, mintPower, mlngClassId, , True)
    If lngRetuId = 0 Then Exit Sub
    
    Call zlSubRefList(lngRetuId)
End Sub

Public Sub zlExecuteControl(ByVal Control As XtremeCommandBars.ICommandBarControl)
    '******************************************************************************************************************
    '功能：执行指定命令条控件，属于维护管理接口
    '参数：指定查找的控件
    '******************************************************************************************************************
    Call cbsThis_Execute(Control)
End Sub

Public Sub zlUpdateControl(ByVal Control As XtremeCommandBars.ICommandBarControl)
    '******************************************************************************************************************
    '功能：刷新命令条控件状态，属于维护管理接口
    '参数：指定查找的控件
    '******************************************************************************************************************
    Call cbsThis_Update(Control)
End Sub


'以下为内部公共程序
'######################################################################################################################
Private Function zlGetPower() As Integer
    '******************************************************************************************************************
    '功能：获得当前用户的词句管理的权限
    '返回：词句管理权限数值
    '******************************************************************************************************************
    If mintPower = con_UnDefine Then
        If InStr(1, gstrPrivsEpr, "全院病历词句") <> 0 Then
            mintPower = 0
        ElseIf InStr(1, gstrPrivsEpr, "科室病历词句") <> 0 Then
            mintPower = 1
        ElseIf InStr(1, gstrPrivsEpr, "个人病历词句") <> 0 Then
            mintPower = 2
        Else
            mintPower = -1
        End If
    End If
    zlGetPower = mintPower
End Function

Private Function zlSubRefClass() As Boolean
    '******************************************************************************************************************
    '功能：刷新分类
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    If mblnCompend = False Then Exit Function
    
    gstrSQL = "Select /*+ rule*/ Id,上级id,编码,名称 From 病历词句分类 Start With Id In ("
    
    
    '------------------------------------------------------------------------------------------------------------------
    strSQL = "Select L.分类id " & vbNewLine & _
            "From 病历词句示范 L, 病历提纲词句 A," & vbNewLine & _
            "     Table(Cast(f_Sentence_Usable([1], [2], [3], [4]) As " & gstrDbOwner & ".t_Dic_Rowset)) U" & vbNewLine & _
            "Where L.分类id = A.词句分类id  And A.提纲id = [1] And L.ID = To_Number(U.编码)"
            
    Select Case mintPower
    Case 0
    Case 1
        strSQL = strSQL & " And" & vbNewLine & _
                "      (Nvl(L.通用级, 0) = 0 Or" & vbNewLine & _
                "      L.通用级 In (1, 2) And" & vbNewLine & _
                "      L.科室id In (Select R.部门id From 部门人员 R, 上机人员表 U Where R.人员id = U.人员id And U.用户名 = User))"

    Case Else
        strSQL = strSQL & " And" & vbNewLine & _
                "      (Nvl(L.通用级, 0) = 0 Or" & vbNewLine & _
                "      L.通用级 = 1 And" & vbNewLine & _
                "      L.科室id In (Select R.部门id From 部门人员 R, 上机人员表 U Where R.人员id = U.人员id And U.用户名 = User) Or" & vbNewLine & _
                "      L.通用级 = 2 And L.人员id In (Select U.人员id From 上机人员表 U Where U.用户名 = User))"
    End Select
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = gstrSQL & strSQL
    gstrSQL = gstrSQL & ") Connect By Prior 上级id=Id  Order By 编码"
    
    Dim objNode As Node
    
    TreeList.Nodes.Clear
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngParentId, mlngPatient, mlngVisit, mlngAdvice)
    If rsTemp.BOF = False Then
                
        Set objNode = TreeList.Nodes.Add(, , "K0", "所有词句", "close", "expend")
        objNode.Expanded = False
        Do While Not rsTemp.EOF
            
            Set objNode = Nothing
            
            On Error Resume Next
            Set objNode = TreeList.Nodes("K" & rsTemp("ID").Value)
            On Error GoTo errHand
            
            If objNode Is Nothing Then
                Set objNode = TreeList.Nodes.Add("K" & zlCommFun.NVL(rsTemp("上级id").Value, 0), tvwChild, "K" & rsTemp("ID").Value, rsTemp("名称").Value, "close", "expend")
                objNode.Expanded = False
            End If
            rsTemp.MoveNext
        Loop
    End If
    If TreeList.Nodes.Count > 0 Then
        TreeList.Nodes(1).Selected = True
    End If
    
    zlSubRefClass = True
    
    Exit Function
errHand:
    
End Function

Private Function zlSubRefList(Optional lngID As Long, Optional ByVal lng分类id As Long) As Long
    '******************************************************************************************************************
    '功能：刷新装入清单，并定位到指定的记录上
    '参数：
    '返回：
    '******************************************************************************************************************
Dim rsTemp As New ADODB.Recordset
Dim strClassIds As String, strKinds As String, strText As String, blnAdd As Boolean
Dim rptRcd As ReportRecord
Dim rptItem As ReportRecordItem
Dim rptRow As ReportRow
    
    '------------------------------------------------------------------------------------------------------------------
    If mblnCompend = False Then
        '按分类显示，属于维护管理程序功能
        
        gstrSQL = "Select /*+ rule*/ L.ID, L.分类id, C.编码 || '-' || C.名称 As 分类, L.编号, L.名称, L.通用级, D.名称 As 部门, P.姓名 As 人员" & vbNewLine & _
                "From 病历词句分类 C, 病历词句示范 L, 部门表 D, 人员表 P" & vbNewLine & _
                "Where C.ID = L.分类id And L.科室id = D.ID And L.人员id = P.ID And L.分类id = [1] "
    Else
        '按提纲显示，属于病历编辑程序功能
        gstrSQL = "Select /*+ rule*/ L.ID, L.分类id, C.编码 || '-' || C.名称 As 分类, L.编号, L.名称, L.通用级, D.名称 As 部门, P.姓名 As 人员" & vbNewLine & _
                "From 病历词句分类 C, 病历词句示范 L, 病历提纲词句 A, 部门表 D, 人员表 P," & vbNewLine & _
                "     Table(Cast(f_Sentence_Usable([1], [2], [3], [4]) As " & gstrDbOwner & ".t_Dic_Rowset)) U" & vbNewLine & _
                "Where C.ID = L.分类id And L.分类id = A.词句分类id And L.科室id = D.ID And L.人员id = P.ID And A.提纲id = [1] And" & vbNewLine & _
                "      L.ID = To_Number(U.编码)  "
        If lng分类id > 0 Then gstrSQL = gstrSQL & "  And L.分类id=[5] "
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    Select Case mintPower
    Case 0
    Case 1
        gstrSQL = gstrSQL & " And" & vbNewLine & _
                "      (Nvl(L.通用级, 0) = 0 Or" & vbNewLine & _
                "      L.通用级 In (1, 2) And" & vbNewLine & _
                "      L.科室id In (Select R.部门id From 部门人员 R, 上机人员表 U Where R.人员id = U.人员id And U.用户名 = User))"

    Case Else
        gstrSQL = gstrSQL & " And" & vbNewLine & _
                "      (Nvl(L.通用级, 0) = 0 Or" & vbNewLine & _
                "      L.通用级 = 1 And" & vbNewLine & _
                "      L.科室id In (Select R.部门id From 部门人员 R, 上机人员表 U Where R.人员id = U.人员id And U.用户名 = User) Or" & vbNewLine & _
                "      L.通用级 = 2 And L.人员id In (Select U.人员id From 上机人员表 U Where U.用户名 = User))"
    End Select
    
    Err = 0: On Error GoTo errHand
    If mblnCompend = False Then
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngParentId)
    Else
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngParentId, mlngPatient, mlngVisit, mlngAdvice, lng分类id)
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    rptList.GroupsOrder.DeleteAll
    rptList.Records.DeleteAll
    strClassIds = ","
    With rsTemp
        Do While Not .EOF
            blnAdd = True
            If mstrSecondLimit <> "" Then '二次过滤
                If InStr(mstrSecondLimit, "," & !ID & ",") = 0 Then blnAdd = False  '不在二次过滤范围内则不加入列表
            End If
            
            If blnAdd Then
                If InStr(1, strClassIds, "," & !分类id & ",") = 0 Then strClassIds = strClassIds & !分类id & ","
                Set rptRcd = Me.rptList.Records.Add()
                Set rptItem = rptRcd.AddItem(CInt(Val("" & !通用级))): rptItem.Icon = rptItem.Value
                Select Case rptItem.Value
                Case 0: rptItem.GroupCaption = "1-全院"
                Case 1: rptItem.GroupCaption = "2-科室"
                Case Else: rptItem.GroupCaption = "3-个人"
                End Select
                rptRcd.AddItem CStr(!ID)
                rptRcd.AddItem CStr("" & !分类id)
                rptRcd.AddItem CStr("" & !分类)
                rptRcd.AddItem CStr("" & !编号)
                rptRcd.AddItem CStr("" & !名称)
                rptRcd.AddItem CStr("" & !部门)
                rptRcd.AddItem CStr("" & !人员)
            End If
            .MoveNext
        Loop
    End With
    
    If mblnCompend = True And UBound(Split(strClassIds, ",")) > 2 Then Me.rptList.GroupsOrder.Add Me.rptList.Columns(mCol.分类)
    Me.rptList.Populate
    
    If lngID <> 0 Then
        For Each rptRow In Me.rptList.Rows
            If rptRow.GroupRow = False Then
                If Val(rptRow.Record(mCol.ID).Value) = lngID Then
                    Set Me.rptList.FocusedRow = rptRow: Exit For
                End If
            End If
        Next
    Else
        For Each rptRow In Me.rptList.Rows
            If rptRow.GroupRow Then rptRow.Expanded = False
        Next
    End If
    If Me.rptList.Rows.Count > 0 And (Me.rptList.FocusedRow Is Nothing) Then
        Set Me.rptList.FocusedRow = Me.rptList.Rows(0)
    End If
    
    Call rptList_SelectionChanged
    zlSubRefList = Me.rptList.Records.Count
    Exit Function

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlSubRefList = Me.rptList.Records.Count
End Function

'以下为控件事件处理
'######################################################################################################################
Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
Dim lngRetuId As Long, strTemp As String
    
    Err = 0: On Error GoTo errHand
    '------------------------------------
    Select Case Control.ID
    Case conMenu_Edit_NewItem
        If Not (Me.rptList.FocusedRow Is Nothing) Then mlngClassId = Me.rptList.FocusedRow.Record(mCol.分类id).Value
        lngRetuId = frmSentenceEdit.ShowMe(mfrmParent, True, mintPower, mlngClassId)
        If lngRetuId = 0 Then Exit Sub
        Call zlSubRefList(lngRetuId)
    Case conMenu_Edit_Modify
        mlngClassId = Me.rptList.FocusedRow.Record(mCol.分类id).Value
        lngRetuId = frmSentenceEdit.ShowMe(mfrmParent, False, mintPower, mlngClassId, mlngWordId)
        If lngRetuId = 0 Then Exit Sub
        Call zlSubRefList(lngRetuId)
    Case conMenu_Edit_Delete
        strTemp = "真的删除该词句吗？" & vbCrLf & "——" & Me.rptList.FocusedRow.Record(mCol.编号).Value & "-" & Me.rptList.FocusedRow.Record(mCol.名称).Value
        If MsgBox(strTemp, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        gstrSQL = "Zl_病历词句示范_Edit(3," & mlngWordId & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        With Me.rptList
            mlngWordId = 0: lngRetuId = .FocusedRow.Index
            If .Rows.Count > lngRetuId + 1 Then
                mlngWordId = .Rows(lngRetuId + 1).Record(mCol.ID).Value
            ElseIf lngRetuId > 0 Then
                mlngWordId = .Rows(lngRetuId - 1).Record(mCol.ID).Value
            End If
        End With
        Call zlSubRefList(mlngWordId)
    Case conMenu_Edit_Request
        If frmSentenceRequest.ShowMe(mfrmParent, mlngWordId) = True Then Call zlSubRefList(mlngWordId)
    Case conMenu_File_Preview: Call zlRptPrint(0)
    Case conMenu_File_Print: Call zlRptPrint(1)
    Case conMenu_File_Excel: Call zlRptPrint(3)
    End Select
    Exit Sub
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Exit Sub
End Sub



Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Err = 0: On Error Resume Next
    Select Case Control.ID
    Case conMenu_Edit_NewItem
        Control.Visible = (mintPower >= 0 And mlngClassId <> 0)
    Case conMenu_Edit_Modify, conMenu_Edit_Delete, conMenu_Edit_Request
        Control.Visible = (mintPower >= 0 And mlngClassId <> 0)
        Control.Enabled = (mlngWordId <> 0)
        If Control.Enabled Then Control.Enabled = (Me.rptList.FocusedRow.Record(mCol.图标).Value >= mintPower)
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Control.Enabled = (Me.rptList.Records.Count <> 0)
    End Select
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    On Error Resume Next
    Select Case Item.ID
    Case conPane_Tree
        Item.Handle = TreeList.hwnd
    Case conPane_List
        Item.Handle = rptList.hwnd
    Case conPane_Term
        Item.Handle = Me.picTerm.hwnd
    Case conPane_Text
        Item.Handle = Me.rtbText.hwnd
    End Select
End Sub

Private Sub Form_Load()
    '-----------------------------------------------------
    '权限限制串复制，避免同时进入其他模块而导致gmstrPrivs变化，导致控制无效
Dim cbrControl As CommandBarControl
Dim cbrMenuBar As CommandBarPopup
Dim rptCol As ReportColumn
    mintPower = con_UnDefine
    mintPower = zlGetPower
    mlngWordId = 0
    
    Call zlCommFun.SetWindowsInTaskBar(Me.hwnd, False)
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbsThis.VisualTheme = xtpThemeOffice2003
    Set Me.cbsThis.Icons = zlCommFun.GetPubIcons
    With Me.cbsThis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbsThis.EnableCustomization False
    
    '-----------------------------------------------------
    '菜单定义
    Me.cbsThis.ActiveMenuBar.Title = "菜单": Me.cbsThis.ActiveMenuBar.Visible = False
    Me.cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)

    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    cbrMenuBar.ID = conMenu_EditPopup
    With cbrMenuBar.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新增(&A)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除(&D)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Request, "限制条件(&Q)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览(&V)"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Excel, "输出到&Excel…")
    End With
    
    '-----------------------------------------------------
    '设置词句显示停靠窗格
    Dim panThis As Pane
    
    Set panThis = dkpMan.CreatePane(conPane_Tree, 600, 300, DockTopOf, Nothing)
    panThis.Title = "树型结构"
    panThis.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Set panThis = dkpMan.CreatePane(conPane_List, 600, 450, DockBottomOf, panThis)
    panThis.Title = "条件列表"
    panThis.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Set panThis = dkpMan.CreatePane(conPane_Text, 600, 800, DockBottomOf, panThis)
    panThis.Title = "示范内容"
    panThis.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    Set panThis = dkpMan.CreatePane(conPane_Term, 200, 800, DockRightOf, Nothing)
    panThis.Title = "示范条件"
    panThis.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    panThis.Close '默认情况下条件内容是不显示的
    
    Me.dkpMan.SetCommandBars Me.cbsThis
    Me.dkpMan.Options.ThemedFloatingFrames = True
    Me.dkpMan.Options.HideClient = True
    dkpMan.LoadStateFromString GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMan), dkpMan.Name, "")
    '-----------------------------------------------------
    With Me.rptList
        Set rptCol = .Columns.Add(mCol.图标, "", 18, False): rptCol.Editable = False: rptCol.Groupable = False
        rptCol.Sortable = False: rptCol.Alignment = xtpAlignmentCenter
        Set rptCol = .Columns.Add(mCol.ID, "ID", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.分类id, "分类id", 0, False): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.分类, "分类", 200, True): rptCol.Editable = False: rptCol.Groupable = False: rptCol.Visible = False
        Set rptCol = .Columns.Add(mCol.编号, "编号", 50, False): rptCol.Editable = False: rptCol.Groupable = False: .SortOrder.Add rptCol
        Set rptCol = .Columns.Add(mCol.名称, "名称", 120, True): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.部门, "部门", 70, False): rptCol.Editable = False: rptCol.Groupable = False
        Set rptCol = .Columns.Add(mCol.人员, "人员", 56, False): rptCol.Editable = False: rptCol.Groupable = False
        .SetImageList Me.imgList
        .AllowColumnRemove = False
        .MultipleSelection = False
        .ShowItemsInGroups = False
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .GridLineColor = RGB(225, 225, 225)
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有可显示的项目..."
            .VerticalGridStyle = xtpGridSolid
        End With
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Set mfrmParent = Nothing
    Call SaveSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMan), dkpMan.Name, dkpMan.SaveStateToString)
End Sub

Private Sub picTerm_Resize()
    Err = 0: On Error Resume Next
    With Me.vfgTerm
        .Left = 0: .Width = Me.picTerm.ScaleWidth
        .Top = 0: .Height = Me.picTerm.ScaleHeight
        .AutoSize 0
    End With
End Sub

Private Sub rptList_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    With Me.rptList
        If .Visible = False Then Exit Sub
        If .FocusedRow Is Nothing Then Exit Sub
        If .FocusedRow.GroupRow Then Exit Sub
        Call rptList_RowDblClick(.FocusedRow, .FocusedRow.Record.Item(mCol.ID))
    End With
End Sub

Private Sub rptList_MouseUp(Button As Integer, Shift As Integer, x As Long, y As Long)
Dim cbrPopupBar As CommandBar
Dim cbrPopupItem As CommandBarControl
Dim cbrControl As CommandBarControl
Dim cbrMenuBar As CommandBarPopup
    
    If Button <> vbRightButton Then Exit Sub
     
    Set cbrMenuBar = Me.cbsThis.ActiveMenuBar.FindControl(xtpControlPopup, conMenu_EditPopup)
    If cbrMenuBar Is Nothing Then Exit Sub
    If cbrMenuBar.Visible = False Then Exit Sub
    
    Set cbrPopupBar = Me.cbsThis.Add("弹出菜单", xtpBarPopup)
    For Each cbrControl In cbrMenuBar.CommandBar.Controls
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, cbrControl.ID, cbrControl.Caption)
        cbrPopupItem.BeginGroup = cbrControl.BeginGroup
    Next
    cbrPopupBar.ShowPopup
End Sub

Private Sub rptList_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
Dim cbrControl As CommandBarControl
    If Me.rptList.FocusedRow Is Nothing Then
        mlngWordId = 0
    ElseIf Me.rptList.FocusedRow.GroupRow = True Then
        mlngWordId = 0
    Else
        mlngWordId = Me.rptList.FocusedRow.Record.Item(mCol.ID).Value
    End If
    If mlngWordId = 0 Then Exit Sub
    
    If mblnCompend = False Then
        If rptList.FocusedRow.Record(mCol.图标).Value >= mintPower Then
            Set cbrControl = Me.cbsThis.ActiveMenuBar.FindControl(xtpControlButton, conMenu_Edit_Modify, True, True)
            If cbrControl Is Nothing Then Exit Sub
            If cbrControl.Visible = False Or cbrControl.Enabled = False Then Exit Sub
            Call cbsThis_Execute(cbrControl)
        End If
    Else
        RaiseEvent RowDblClick(mlngWordId)
    End If
End Sub

Private Sub rptList_SelectionChanged()
    Dim rsTemp As New ADODB.Recordset
    Dim lngStart As Long, strText As String
    
    If Me.rptList.FocusedRow Is Nothing Then
        mlngWordId = 0
    ElseIf Me.rptList.FocusedRow.GroupRow = True Then
        mlngWordId = 0
    Else
        mlngWordId = Me.rptList.FocusedRow.Record.Item(mCol.ID).Value
    End If
    If Me.Visible = False Then Exit Sub

    '刷新词句内容
    '------------------------------------------------------------------------------------------------------------------
    Me.rtbText.Text = ""
    Err = 0: On Error GoTo errHand
    gstrSQL = "Select 内容性质, 内容文本, 要素名称, 要素单位 From 病历词句组成 Where 词句id = [1] Order By 排列次序"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngWordId)
    With rsTemp
        Do While Not .EOF
            lngStart = Len(Me.rtbText.Text)
            Me.rtbText.SelStart = lngStart
            Me.rtbText.SelLength = 0
            Select Case !内容性质
            Case 0 '自由文字
                strText = IIf(IsNull(!内容文本), " ", !内容文本)
                With Me.rtbText
                    .SelText = strText: .SelStart = lngStart: .SelLength = Len(strText)
                    .SelUnderline = False
                End With
            Case 1, 2 '1-临时诊治要素,2-固定诊治要素
                strText = IIf(IsNull(!内容文本), "{" & !要素名称 & "}" & !要素单位, "{" & !内容文本 & "}")
                With Me.rtbText
                    .SelText = strText: .SelStart = lngStart: .SelLength = Len(strText)
                    .SelUnderline = True
                End With
            End Select
            .MoveNext
        Loop
        Me.rtbText.SelStart = 0
    End With
    
    '刷新词句条件
    Dim panThis As Pane
    Set panThis = Me.dkpMan.FindPane(conPane_Term)
    If panThis Is Nothing Then Exit Sub
    If panThis.Closed Then Exit Sub
    
    Me.vfgTerm.Clear: Me.vfgTerm.Rows = Me.vfgTerm.FixedRows
    Set Me.vfgTerm.Cell(flexcpPicture, Me.vfgTerm.FixedRows - 1, 0) = Me.imgList.ListImages(4).Picture
    gstrSQL = "Select 名称 As 条件项, 简码 As 条件值" & vbNewLine & _
            "From Table(Cast(f_Sentence_条件项([1]) As " & gstrDbOwner & ".t_Dic_Rowset))" & vbNewLine & _
            "Where 简码 Is Not Null"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngWordId)
    With rsTemp
        If .RecordCount <= 0 Then
            Me.vfgTerm.TextMatrix(Me.vfgTerm.FixedRows - 1, 0) = "无使用限制条件。"
        Else
            Me.vfgTerm.TextMatrix(Me.vfgTerm.FixedRows - 1, 0) = "在以下条件满足时可以使用："
        End If
        Do While Not .EOF
            Me.vfgTerm.Rows = Me.vfgTerm.Rows + 1
            Me.vfgTerm.TextMatrix(Me.vfgTerm.Rows - 1, 0) = Space(2) & Me.vfgTerm.Rows - 1 & ")" & !条件项 & "为'" & Replace(!条件值, vbTab, "'或'") & "'"
            .MoveNext
        Loop
    End With
    Me.vfgTerm.AutoSize 0
    
    Exit Sub
    
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub TreeList_NodeClick(ByVal Node As MSComctlLib.Node)
    Call zlSubRefList(mlngWordId, Val(Mid(Node.Key, 2)))
End Sub
Private Sub zlRptPrint(ByVal bytMode As Byte)
    '******************************************************************************************************************
    '功能:将数据复制到可打印的对象，调用打印
    '参数:  bytMode，1-打印;2-预览;3-输出到EXCEL
    '******************************************************************************************************************
    
    Dim objPrint As New zlPrint1Grd
    Dim objAppRow As zlTabAppRow
    
    If Me.rptList.Records.Count = 0 Then Exit Sub
    
    '-------------------------------------------------
    '复制数据表格
    If zlReportToVSFlexGrid(Me.vfgListTmp, Me.rptList) = False Then Exit Sub
    
    '-------------------------------------------------
    '调用打印部件处理

    
    Set objPrint.Body = Me.vfgListTmp
    objPrint.Title.Text = "词句示范清单"
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objAppRow.Add("打印时间:" & Now())
    Call objPrint.BelowAppRows.Add(objAppRow)
    
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub
