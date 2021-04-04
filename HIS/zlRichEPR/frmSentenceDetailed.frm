VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.Unicode.9600.ocx"
Begin VB.Form frmSentenceDetailed 
   BorderStyle     =   0  'None
   Caption         =   "词句过滤"
   ClientHeight    =   7290
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6345
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7290
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox piclist 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   600
      ScaleHeight     =   1935
      ScaleWidth      =   2655
      TabIndex        =   4
      Top             =   4680
      Width           =   2655
      Begin VSFlex8Ctl.VSFlexGrid vsfList 
         Height          =   1695
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   1695
         _cx             =   2990
         _cy             =   2990
         Appearance      =   1
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
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   16777215
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
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
   End
   Begin VB.PictureBox picfind 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   3240
      ScaleHeight     =   1335
      ScaleWidth      =   2655
      TabIndex        =   0
      Top             =   1080
      Width           =   2655
      Begin VB.TextBox txtFind 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   840
         TabIndex        =   6
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label lblFilter 
         BackColor       =   &H00FFFFFF&
         Caption         =   "过滤："
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   615
      End
      Begin VB.Shape shpFind 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00400040&
         Height          =   255
         Left            =   0
         Top             =   120
         Width           =   1215
      End
   End
   Begin MSComctlLib.ImageList imgClass 
      Left            =   1800
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
            Picture         =   "frmSentenceDetailed.frx":0000
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSentenceDetailed.frx":059A
            Key             =   "expend"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TreeList 
      Height          =   1470
      Left            =   240
      TabIndex        =   2
      Top             =   720
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
   Begin MSComctlLib.ImageList imgList 
      Left            =   720
      Top             =   2280
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
            Picture         =   "frmSentenceDetailed.frx":0B34
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSentenceDetailed.frx":10CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSentenceDetailed.frx":1668
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSentenceDetailed.frx":1C02
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfFind 
      Height          =   615
      Left            =   600
      TabIndex        =   3
      Top             =   3720
      Visible         =   0   'False
      Width           =   2175
      _cx             =   3836
      _cy             =   1085
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16777215
      BackColorAlternate=   16777215
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   0
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
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   360
      Top             =   240
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmSentenceDetailed.frx":24DC
      Left            =   840
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmSentenceDetailed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'######################################################################################################################

'常量定义
'------------------------------------------------------------------------------------------------------------------
Private Enum mCol
    ID = 0: sortid: Sort: Num: pName: Range: Depart: personnel: Pinyin: Wubi
End Enum


Private Const conPane_Tree = 400
Private Const conPane_List = 401
Private Const conPane_Text = 404
Private cbrControl As CommandBarControl
Private cbrMenuBar As CommandBarPopup
Private cbrToolBar As CommandBar

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
Private mfrmTipInfo As New frmTipInfo
Private mlngId As Long
Private mrsTmp As ADODB.Recordset
Private mintPrompt As Integer '判断是否刷新提示的信息
Private mstrContent As String
Private mintPower As Integer
Private mLeftRight As Integer
Public Event RowDblClick(ByVal lngSentenceID As Long)    '双击一行或在行上按回车
Public Event ShiftFocus()           '改变焦点


'以下为外部公共程序
'######################################################################################################################

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
    mlngClassId = lngPatient
    Set mfrmParent = frmParent
    If blnForce = False And mlngParentId = lngCompendID And _
        (mstrSecondLimit = strSecondLimit) Then zlRefFromCompend = Me.vsfList.Rows: Exit Function
    mblnCompend = True
    TreeList.Visible = mblnCompend
    mlngParentId = lngCompendID
    mlngPatient = lngPatient
    mlngVisit = lngVisit
    mlngAdvice = lngAdvice
    mstrSecondLimit = strSecondLimit
    
    Set panThis = dkpMan.FindPane(conPane_Tree)
    panThis.Closed = False
    
    gstrSQL = "Select 词句分类id From 病历提纲词句 Where 提纲id = [1]"
    Err = 0: On Error GoTo errHand
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "frmSentenceDetailed", mlngParentId)
  
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



'以下为内部公共程序
'######################################################################################################################

Public Function zlSubRefList(Optional lngID As Long, Optional ByVal lng分类id As Long) As Long
    '******************************************************************************************************************
    '功能：刷新装入清单，并定位到指定的记录上
    '参数：
    '返回：
    '******************************************************************************************************************
Dim rsTemp As New ADODB.Recordset
Dim strClassIds As String, strKinds As String, strText As String, blnAdd As Boolean
Dim i As Integer


    
    '------------------------------------------------------------------------------------------------------------------
        '按提纲显示，属于病历编辑程序功能
        gstrSQL = "Select /*+ rule*/ L.ID, L.分类id, C.编码 || '-' || C.名称 As 分类, L.编号, L.名称, L.通用级 as 范围, D.名称 As 部门, P.姓名 As 人员,zlspellcode(L.名称) as 拼音,zlwbcode(L.名称) as 五笔" & vbNewLine & _
                "From 病历词句分类 C, 病历词句示范 L, 病历提纲词句 A, 部门表 D, 人员表 P," & vbNewLine & _
                "     Table(Cast(f_Sentence_Usable([1], [2], [3], [4]) As " & gstrDbOwner & ".t_Dic_Rowset)) U" & vbNewLine & _
                "Where C.ID = L.分类id And L.分类id = A.词句分类id And L.科室id = D.ID And L.人员id = P.ID And A.提纲id = [1] And" & vbNewLine & _
                "      L.ID = To_Number(U.编码)  "
        If lng分类id > 0 Then gstrSQL = gstrSQL & "  And L.分类id=[5] "

    '------------------------------------------------------------------------------------------------------------------
    If InStr(1, gstrPrivsEpr, "科室病历词句") <> 0 Then
        gstrSQL = gstrSQL & " And" & vbNewLine & _
                "      (Nvl(L.通用级, 0) = 0 Or" & vbNewLine & _
                "      L.通用级 In (1, 2) And" & vbNewLine & _
                "      L.科室id In (Select R.部门id From 部门人员 R, 上机人员表 U Where R.人员id = U.人员id And U.用户名 = User))"

     Else
        gstrSQL = gstrSQL & " And" & vbNewLine & _
                "      (Nvl(L.通用级, 0) = 0 Or" & vbNewLine & _
                "      L.通用级 = 1 And" & vbNewLine & _
                "      L.科室id In (Select R.部门id From 部门人员 R, 上机人员表 U Where R.人员id = U.人员id And U.用户名 = User) Or" & vbNewLine & _
                "      L.通用级 = 2 And L.人员id In (Select U.人员id From 上机人员表 U Where U.用户名 = User))"
    End If
    
    Err = 0: On Error GoTo errHand
    gstrSQL = gstrSQL & "Order by L.通用级 Desc, Lpad(L.编号,13,'0')"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "frmSentenceDetailed", mlngParentId, mlngPatient, mlngVisit, mlngAdvice, lng分类id)
     Set mrsTmp = rsTemp
    '------------------------------------------------------------------------------------------------------------------

    strClassIds = ","

    If Not rsTemp.EOF Then
        With Me.vsfList
            Set .DataSource = rsTemp
            .ColWidth(mCol.ID) = 0
            .ColWidth(mCol.sortid) = 0
            .ColWidth(mCol.Sort) = 0
            .ColWidth(mCol.Num) = Me.picList.Width / 6 + 200
            .ColWidth(mCol.pName) = Me.picList.Width / 3 * 2 - 200
            .ColWidth(mCol.Range) = Me.picList.Width / 6 - 50
            .ColWidth(mCol.Depart) = 0
            .ColWidth(mCol.personnel) = 0
            .ColWidth(mCol.Pinyin) = 0
            .ColWidth(mCol.Wubi) = 0
            For i = 1 To .Rows - 1
                Select Case .TextMatrix(i, mCol.Range)
                    Case 0:
                        .TextMatrix(i, 5) = ""
                        .Cell(flexcpPicture, i, mCol.Range) = Me.imgList.ListImages(1).Picture '"1-全院"
                    Case 1:
                        .TextMatrix(i, 5) = ""
                        .Cell(flexcpPicture, i, mCol.Range) = Me.imgList.ListImages(2).Picture '"2-科室"
                    Case 2:
                        .TextMatrix(i, 5) = ""
                        .Cell(flexcpPicture, i, mCol.Range) = Me.imgList.ListImages(3).Picture '"3-个人"
                    Case Else:
                        .TextMatrix(i, 5) = ""
                        .Cell(flexcpPicture, i, mCol.Range) = Me.imgList.ListImages(1).Picture '"1-全院"
                End Select
            Next
        End With
    Else
        Me.vsfList.Rows = 1
    End If

    If mlngId <> mlngParentId Or rsTemp.RecordCount = 0 Then
        Me.vsfFind.Visible = False
        Me.txtFind.Text = ""
        mlngId = mlngParentId
    End If
    If vsfList.Rows > 1 Then
        vsfList.Row = 1
    End If
    Exit Function
    
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlSubRefList = Me.vsfList.Rows
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
            
    If InStr(1, gstrPrivsEpr, "科室病历词句") <> 0 Then
        strSQL = strSQL & " And" & vbNewLine & _
                "      (Nvl(L.通用级, 0) = 0 Or" & vbNewLine & _
                "      L.通用级 In (1, 2) And" & vbNewLine & _
                "      L.科室id In (Select R.部门id From 部门人员 R, 上机人员表 U Where R.人员id = U.人员id And U.用户名 = User))"

    Else
        strSQL = strSQL & " And" & vbNewLine & _
                "      (Nvl(L.通用级, 0) = 0 Or" & vbNewLine & _
                "      L.通用级 = 1 And" & vbNewLine & _
                "      L.科室id In (Select R.部门id From 部门人员 R, 上机人员表 U Where R.人员id = U.人员id And U.用户名 = User) Or" & vbNewLine & _
                "      L.通用级 = 2 And L.人员id In (Select U.人员id From 上机人员表 U Where U.用户名 = User))"
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = gstrSQL & strSQL
    gstrSQL = gstrSQL & ") Connect By Prior 上级id=Id  Order By 编码"
    
    Dim objNode As node
    
    TreeList.Nodes.Clear
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "frmSentenceDetailed", mlngParentId, mlngPatient, mlngVisit, mlngAdvice)
    If rsTemp.BOF = False Then
                
        Set objNode = TreeList.Nodes.Add(, , "K0", "所有词句", "close", "expend")
        objNode.Expanded = True
        Do While Not rsTemp.EOF
            
            Set objNode = Nothing
            
            On Error Resume Next
            Set objNode = TreeList.Nodes("K" & rsTemp("ID").Value)
            objNode.Expanded = True
            On Error GoTo errHand
            
            If objNode Is Nothing Then
                Set objNode = TreeList.Nodes.Add("K" & zlCommFun.NVL(rsTemp("上级id").Value, 0), tvwChild, "K" & rsTemp("ID").Value, rsTemp("名称").Value, "close", "expend")
                objNode.Expanded = True
            End If
            rsTemp.MoveNext
        Loop
    End If
    If TreeList.Nodes.Count > 0 Then
        TreeList.Nodes(1).Selected = True
    Else
        mlngClassId = 0
    End If
    zlSubRefClass = True
    
    Exit Function
errHand:
    
End Function

'以下为控件事件处理
'######################################################################################################################
Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    On Error Resume Next
    Select Case Item.ID
    Case conPane_Tree
        Item.Handle = TreeList.hwnd
    Case conPane_List
        Item.Handle = picList.hwnd
    Case conPane_Text
        Item.Handle = Me.picfind.hwnd
    End Select
End Sub


Private Sub Form_Load()
    '-----------------------------------------------------
    '权限限制串复制，避免同时进入其他模块而导致gmstrPrivs变化，导致控制无效
Dim rptCol As ReportColumn
     mlngWordId = 0
     
     Call zlCommFun.SetWindowsInTaskBar(Me.hwnd, False)
    
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
     
     Set panThis = dkpMan.CreatePane(conPane_Text, 600, 50, DockTopOf, Nothing)
     panThis.Title = "快捷查询"
     panThis.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
     panThis.MaxTrackSize.Height = 23
     panThis.MinTrackSize.Height = 23
     
     Set panThis = dkpMan.CreatePane(conPane_Tree, 600, 300, DockBottomOf, panThis)
     panThis.Title = "树型结构"
     panThis.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
     
     Set panThis = dkpMan.CreatePane(conPane_List, 600, 450, DockBottomOf, panThis)
     panThis.Title = "条件列表"
     panThis.Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
     panThis.Selected = False
    
    
     Me.dkpMan.Options.ThemedFloatingFrames = True
     Me.dkpMan.Options.HideClient = True
     dkpMan.LoadStateFromString GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & "frmSentenceDetailed" & "\" & TypeName(dkpMan), dkpMan.Name, "")
     '-----------------------------------------------------
    With Me.vsfFind
         .FixedCols = 0
         .SelectionMode = flexSelectionByRow
    End With
    '对vsflist进行初始话
    With Me.vsfList
         .SelectionMode = flexSelectionByRow
         .FixedCols = 0
         .ExplorerBar = flexExSortShow
         .AddItem ""
         .Rows = 1
         .Cols = 10
         .TextMatrix(0, mCol.ID) = ""
         .TextMatrix(0, mCol.Num) = "编号"
         .TextMatrix(0, mCol.pName) = "名称"
         .TextMatrix(0, mCol.Range) = "范围"
    End With
    If InStr(1, gstrPrivsEpr, "全院病历词句") <> 0 Then
        mintPower = 0
    ElseIf InStr(1, gstrPrivsEpr, "科室病历词句") <> 0 Then
        mintPower = 1
    ElseIf InStr(1, gstrPrivsEpr, "个人病历词句") <> 0 Then
        mintPower = 2
    Else
        mintPower = -1
    End If
End Sub




Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Set cbrControl = Nothing
    Set cbrMenuBar = Nothing
    Set cbrToolBar = Nothing
    Set mfrmParent = Nothing
    Unload mfrmTipInfo
    Set mrsTmp = Nothing
    imgClass.ListImages.Clear
    imgList.ListImages.Clear
    ImageList_Destroy imgClass.hImageList
    ImageList_Destroy imgList.hImageList
    Call SaveSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & "frmSentenceDetailed" & "\" & TypeName(dkpMan), dkpMan.Name, dkpMan.SaveStateToString)
End Sub

Private Sub picfind_Resize()
    On Error Resume Next
    Me.picfind.BackColor = RGB(216, 231, 252)
    Me.lblFilter.BackColor = RGB(216, 231, 252)
    Me.lblFilter.Move 0, 80, Me.picfind.Width / 5, 220
    Me.txtFind.Move Me.lblFilter.Width + Screen.TwipsPerPixelX, 80, Me.picfind.Width / 5 * 4 - 2 * Screen.TwipsPerPixelX, 220
    Me.shpFind.Move Me.lblFilter.Width, 80 - Screen.TwipsPerPixelY, Me.txtFind.Width + 2 * Screen.TwipsPerPixelX, Me.txtFind.Height + 2 * Screen.TwipsPerPixelY
End Sub



Private Function getSentenceContent(lid As Long) As String
    Dim rsTemp As New ADODB.Recordset
    Dim strContent As String
    Dim lngStart As Long
        mlngWordId = lid
    If Me.Visible = False Then Exit Function

    '刷新词句内容
    '------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error GoTo errHand
    gstrSQL = "Select 内容性质, 内容文本, 要素名称, 要素单位 From 病历词句组成 Where 词句id = [1] Order By 排列次序"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "frmSentenceDetailed", mlngWordId)
    With rsTemp
       Do While Not .EOF
            Select Case !内容性质
            Case 0 '自由文字
                strContent = strContent & IIf(IsNull(!内容文本), " ", !内容文本)
            Case 1, 2 '1-临时诊治要素,2-固定诊治要素
                strContent = strContent & IIf(IsNull(!内容文本), "{" & !要素名称 & "}" & !要素单位, "{" & !内容文本 & "}")
            End Select
            .MoveNext
        Loop
    getSentenceContent = strContent
    End With
    Exit Function
    '------------------------------------------------------------------------------------------------------------------
errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub piclist_Resize()
    On Error Resume Next
    Me.vsfList.Move 0, 0, Me.picList.Width, Me.picList.Height
    With Me.vsfList
        .ColWidth(mCol.ID) = 0
        .ColWidth(mCol.sortid) = 0
        .ColWidth(mCol.Sort) = 0
        .ColWidth(mCol.Num) = Me.picList.Width / 6 + 200
        .ColWidth(mCol.pName) = Me.picList.Width / 3 * 2 - 200
        .ColWidth(mCol.Range) = Me.picList.Width / 6 - 50
        .ColWidth(mCol.Depart) = 0
        .ColWidth(mCol.personnel) = 0
        .ColWidth(mCol.Pinyin) = 0
        .ColWidth(mCol.Wubi) = 0
    End With
End Sub

Private Sub TreeList_NodeClick(ByVal node As MSComctlLib.node)
    If Val(Mid(node.Key, 2)) <> 0 Then mlngClassId = Val(Mid(node.Key, 2))
    Call zlSubRefList(mlngWordId, Val(Mid(node.Key, 2)))
End Sub

Private Sub txtFind_Change()
'如果当前词句列表没有数据就不处理
 If Me.vsfList.Rows < 2 Then Exit Sub
    If Me.txtFind.Text <> "" Then
        mrsTmp.Filter = ""
        On Error GoTo aa
        mrsTmp.Filter = "名称 like '*" & Me.txtFind.Text & "*' or 编号 like '*" & Me.txtFind.Text & "*' or 拼音 like '*" & Me.txtFind.Text & "*' or 五笔 like '*" & Me.txtFind.Text & "*'"
        If mrsTmp.RecordCount < 1 Then
            Me.vsfFind.Visible = False
            Exit Sub
        End If
        Set vsfFind.DataSource = mrsTmp
        Me.vsfFind.Move Me.txtFind.Left, Me.picfind.Height, Me.txtFind.Width, (mrsTmp.RecordCount + 2) * Me.vsfFind.ROWHEIGHT(1)
        Me.vsfFind.SheetBorder = RGB(216, 231, 252)
        Me.vsfFind.ZOrder 0
        Me.vsfFind.Visible = True
    Else
        Me.vsfFind.ZOrder 1
        Me.vsfFind.Visible = False
    End If
    With Me.vsfFind
       .ColWidth(0) = 0
       .ColWidth(1) = 0
       .ColWidth(2) = 0
       .ColWidth(3) = Me.vsfFind.Width / 3 - 50
       .ColWidth(4) = Me.vsfFind.Width / 3 * 2
       .ColWidth(5) = 0
       .ColWidth(6) = 0
       .ColWidth(7) = 0
       .ColWidth(8) = 0
       .ColWidth(9) = 0
    End With
    Exit Sub
aa:
    MsgBox "您输入的数据不合法，您只能输入字符和数字以及中文标点符号！", vbInformation, gstrSysName
End Sub

Private Sub txtFind_GotFocus()
    RaiseEvent ShiftFocus
End Sub

Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer

    '按退出按钮
    If KeyCode = vbKeyEscape Then
        Me.vsfFind.Visible = False
        Me.txtFind.Text = ""
        Exit Sub
    End If
    '当按下回车的时候
    If KeyCode = vbKeyReturn Then
        If mrsTmp Is Nothing Then Exit Sub
        '如果记录集没有数据，那么重新查询数据
        If mrsTmp.RecordCount < 1 Then
            Call zlSubRefList(mlngWordId, 0)
            Call txtFind_Change
        Else
            With Me.vsfList
                If .Rows < 2 Then Exit Sub
                If .TextMatrix(.Row, 0) = "" Then Exit Sub
                '选中选择的数据
                For i = 1 To .Rows - 1
                    If .TextMatrix(i, 0) = Me.vsfFind.TextMatrix(Me.vsfFind.Row, 0) Then
                        .Row = i
                    End If
                Next
                '把词句添加到文件中
                RaiseEvent RowDblClick(Me.vsfFind.TextMatrix(Me.vsfFind.Row, 0))
                Me.txtFind.Text = ""
            End With
            Me.vsfFind.Visible = False
        End If
    End If

    
    '当按下下键的时候改变相应的vsf选择行
    If KeyCode = vbKeyDown And Me.vsfFind.Row < Me.vsfFind.Rows - 1 Then
        Me.vsfFind.Row = Me.vsfFind.Row + 1
    End If
    
    '当按下上键的时候改变相应的vsf选择行
    If KeyCode = vbKeyUp And Me.vsfFind.Row > 1 Then
        Me.vsfFind.Row = Me.vsfFind.Row - 1
    End If
End Sub

Private Sub vsfFind_DblClick()
   Call txtFind_KeyDown(vbKeyReturn, -1)
End Sub
Private Sub vsfFind_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
         Me.picfind.Visible = False
    End If
    If KeyCode = vbKeyReturn Then
         Call vsfFind_DblClick
    End If
End Sub

Private Sub vsfList_DblClick()
Dim introw As Integer
    introw = Me.vsfList.MouseRow
    If introw < 1 Then Exit Sub
    mlngWordId = Val(Me.vsfList.TextMatrix(Me.vsfList.Row, mCol.ID))
   
    If mlngWordId = 0 Then Exit Sub
    RaiseEvent RowDblClick(mlngWordId)

End Sub

Private Sub vsfList_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    With Me.vsfList
        If .Rows < 2 Then Exit Sub
        Call vsfList_DblClick
    End With
End Sub


Private Sub vsfList_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim intpro As Integer
    intpro = y \ Me.vsfList.ROWHEIGHT(0)
    If intpro > vsfList.Rows - 1 Then
        Call mfrmTipInfo.ShowTipInfo(vsfList.hwnd, "", True)
        Exit Sub
    End If
    If Me.Width / 6 * 5 < x And x < Me.Width And intpro <> 0 Then
        
        '如果是同一行就不从新刷新词句
        If mintPrompt <> intpro Then
            mstrContent = getSentenceContent(Me.vsfList.TextMatrix(intpro, mCol.ID))
            mintPrompt = intpro
        End If
        Call mfrmTipInfo.ShowTipInfo(vsfList.hwnd, mstrContent, True)
    Else
        Call mfrmTipInfo.ShowTipInfo(vsfList.hwnd, "", True)
   End If
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
Dim lngRetuId As Long, strTemp As String
    
    Err = 0: On Error GoTo errHand
    '------------------------------------
    Select Case Control.ID
    Case conMenu_Edit_NewItem
        If Me.vsfList.Rows > 1 Then mlngClassId = Me.vsfList.TextMatrix(vsfList.Row, mCol.sortid)
        lngRetuId = frmSentenceEdit.ShowMe(mfrmParent, True, mintPower, mlngClassId)
        If lngRetuId = 0 Then Exit Sub
        Call zlSubRefList(lngRetuId)
    Case conMenu_Edit_Modify
        lngRetuId = frmSentenceEdit.ShowMe(mfrmParent, False, mintPower, Me.vsfList.TextMatrix(vsfList.Row, mCol.sortid), Me.vsfList.TextMatrix(vsfList.Row, mCol.ID))
        If lngRetuId = 0 Then Exit Sub
        Call zlSubRefList(lngRetuId)
    Case conMenu_Edit_Delete
        strTemp = "真的删除该词句吗？" & vbCrLf & "——" & Me.vsfList.TextMatrix(vsfList.Row, mCol.Num) & "-" & Me.vsfList.TextMatrix(vsfList.Row, mCol.pName)
        If MsgBox(strTemp, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        gstrSQL = "Zl_病历词句示范_Edit(3," & vsfList.TextMatrix(vsfList.Row, mCol.ID) & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, "frmSentenceDetailed")
        Call zlSubRefList(mlngWordId)
    Case conMenu_Edit_Request
        If frmSentenceRequest.ShowMe(mfrmParent, Me.vsfList.TextMatrix(vsfList.Row, mCol.ID)) = True Then Call zlSubRefList(Me.vsfList.TextMatrix(vsfList.Row, mCol.ID))
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
    Dim lngEnable As Long
    Select Case Control.ID
    Case conMenu_Edit_NewItem
        Control.Visible = (mintPower >= 0 And mlngClassId <> 0)
    Case conMenu_Edit_Modify, conMenu_Edit_Delete, conMenu_Edit_Request
        Control.Visible = (mintPower >= 0 And mlngClassId <> 0)
        If vsfList.Rows > 1 Then
            Control.Enabled = (Me.vsfList.TextMatrix(vsfList.Row, mCol.ID) <> 0)
        Else
            Control.Enabled = False
        End If
        With Me.vsfList
              Select Case .Cell(flexcpPicture, vsfList.Row, mCol.Range)
                  Case Me.imgList.ListImages(1).Picture: '"1-全院"
                       lngEnable = 0
                  Case Me.imgList.ListImages(2).Picture: '"2-科室"
                       lngEnable = 1
                  Case Me.imgList.ListImages(3).Picture: '"3-个人"
                       lngEnable = 2
                  Case Else:
                      lngEnable = 0
              End Select
          End With
        If Control.Enabled Then Control.Enabled = (lngEnable >= mintPower)
        If mintPower = -1 Then Control.Enabled = False
    Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
        Control.Enabled = (Me.vsfList.Rows > 1)
    End Select
End Sub
Private Sub zlRptPrint(ByVal bytMode As Byte)
    '******************************************************************************************************************
    '功能:将数据复制到可打印的对象，调用打印
    '参数:  bytMode，1-打印;2-预览;3-输出到EXCEL
    '******************************************************************************************************************
    
    Dim objPrint As New zlPrint1Grd
    Dim objAppRow As zlTabAppRow
    
    If Me.vsfList.Rows < 1 Then Exit Sub

    
    '-------------------------------------------------
    '调用打印部件处理

    Set objPrint.Body = Me.vsfList
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

Private Sub vsfList_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl

    If Button <> vbRightButton Then Exit Sub

    Set cbrMenuBar = Nothing
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
