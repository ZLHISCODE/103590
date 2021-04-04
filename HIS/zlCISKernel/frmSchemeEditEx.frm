VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSchemeEditEx 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   3150
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4140
   ControlBox      =   0   'False
   Icon            =   "frmSchemeEditEx.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   4140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame fraMethod 
      BackColor       =   &H8000000E&
      Height          =   2175
      Left            =   1560
      TabIndex        =   13
      Top             =   720
      Visible         =   0   'False
      Width           =   2055
      Begin VB.CommandButton cmdMethodOK 
         Caption         =   "确定"
         Height          =   300
         Left            =   1065
         TabIndex        =   14
         Top             =   1800
         Width           =   975
      End
      Begin VSFlex8Ctl.VSFlexGrid vsMethod 
         Height          =   1815
         Left            =   0
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   0
         Width           =   2055
         _cx             =   1993543209
         _cy             =   1993542785
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
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   4210752
         ForeColorSel    =   16777215
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   0
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   0
         Cols            =   2
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   250
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmSchemeEditEx.frx":000C
         ScrollTrack     =   -1  'True
         ScrollBars      =   2
         ScrollTips      =   0   'False
         MergeCells      =   1
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
   End
   Begin VB.Frame fraBorder 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   45
      Index           =   0
      Left            =   495
      MousePointer    =   7  'Size N S
      TabIndex        =   11
      Top             =   2310
      Width           =   615
   End
   Begin VB.Frame fraBorder 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   45
      Index           =   2
      Left            =   495
      TabIndex        =   10
      Top             =   2580
      Width           =   615
   End
   Begin VB.Frame fraBorder 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   330
      Index           =   3
      Left            =   405
      TabIndex        =   9
      Top             =   2295
      Width           =   45
   End
   Begin VB.Frame fraBorder 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00404040&
      Height          =   330
      Index           =   1
      Left            =   1155
      MousePointer    =   9  'Size W E
      TabIndex        =   8
      Top             =   2310
      Width           =   45
   End
   Begin VB.CommandButton cmdData 
      Caption         =   "…"
      Height          =   240
      Left            =   2475
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "选择项目(*)"
      Top             =   1950
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.TextBox txtData 
      Height          =   300
      Left            =   525
      TabIndex        =   3
      Top             =   1920
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.CommandButton cmdCancel 
      Height          =   315
      Left            =   3555
      Picture         =   "frmSchemeEditEx.frx":0048
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "取消(Esc)"
      Top             =   1920
      Width           =   450
   End
   Begin VSFlex8Ctl.VSFlexGrid vsExt 
      Height          =   1845
      Left            =   45
      TabIndex        =   0
      Top             =   30
      Width           =   4080
      _cx             =   7197
      _cy             =   3254
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
      BackColorSel    =   4210752
      ForeColorSel    =   16777215
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   3
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   7
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmSchemeEditEx.frx":05D2
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
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
      Begin MSComctlLib.ImageList img16 
         Left            =   1650
         Top             =   975
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
               Picture         =   "frmSchemeEditEx.frx":06CD
               Key             =   "c0"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSchemeEditEx.frx":0C67
               Key             =   "c1"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSchemeEditEx.frx":1201
               Key             =   "o0"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmSchemeEditEx.frx":179B
               Key             =   "o1"
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmd 
         Caption         =   "…"
         Height          =   240
         Left            =   3435
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "选择项目(*)"
         Top             =   1035
         Visible         =   0   'False
         Width           =   270
      End
   End
   Begin VB.ComboBox cbo标本 
      Height          =   300
      Left            =   525
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1920
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.CommandButton cmdOK 
      Height          =   315
      Left            =   3015
      Picture         =   "frmSchemeEditEx.frx":1D35
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "确认(F2)"
      Top             =   1920
      Width           =   450
   End
   Begin VB.Label Label1 
      Caption         =   "共___味"
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Line lin 
      Index           =   0
      X1              =   2070
      X2              =   2745
      Y1              =   2385
      Y2              =   2385
   End
   Begin VB.Line lin 
      Index           =   1
      X1              =   2070
      X2              =   2745
      Y1              =   2415
      Y2              =   2415
   End
   Begin VB.Line lin 
      Index           =   2
      X1              =   2070
      X2              =   2745
      Y1              =   2445
      Y2              =   2445
   End
   Begin VB.Line lin 
      Index           =   3
      X1              =   2070
      X2              =   2745
      Y1              =   2475
      Y2              =   2475
   End
   Begin VB.Line lin 
      Index           =   4
      X1              =   2070
      X2              =   2745
      Y1              =   2505
      Y2              =   2505
   End
   Begin VB.Line lin 
      Index           =   5
      X1              =   2070
      X2              =   2745
      Y1              =   2535
      Y2              =   2535
   End
   Begin VB.Line lin 
      Index           =   6
      X1              =   2070
      X2              =   2745
      Y1              =   2565
      Y2              =   2565
   End
   Begin VB.Line lin 
      Index           =   7
      X1              =   2070
      X2              =   2745
      Y1              =   2595
      Y2              =   2595
   End
   Begin VB.Label lblData 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "麻醉"
      Height          =   180
      Left            =   105
      TabIndex        =   2
      Top             =   1980
      Visible         =   0   'False
      Width           =   360
   End
End
Attribute VB_Name = "frmSchemeEditEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'入口参数：
Private mlngHwnd As Long '用于定位的控件句柄
Private mint服务对象 As Integer '1-门诊,2-住院,3-门诊和住院
Private mint期效 As Integer

'0-检查组合,1-手术输入,4-检验组合
Private mintType As Integer

'入:主诊疗项目ID
Private mlng项目ID As Long

'入/出:附加定义数据,新增时一般为空
'      检查="部位名1;方法名1,方法名2|部位名2;方法名1,方法名2|...<vbTab>0-常规/1-床旁/2-术中"
'      手术="手术ID1,手术ID2,...;麻醉ID",其中可能没有附加手术和麻醉
'      检验组合="项目ID1,项目ID2,...;检验标本" 如果是新版LIS的模式则是："项目ID1|指标1|指标2...,项目ID2|指标1|指标2...,...;检验标本"
Private mstrExtData As String
Private mblnNew As Boolean  '判断是否是新开输入项目时进入，否则为点下箭头进入
'入：判断检验组合是否使用新版LIS的检验组合模式
Private mblnNewLIS As Boolean
'出口参数：
Private mblnOK As Boolean '出


Private mblnReturn As Boolean '是否了回车确认
Private mblnNotAddNew As Boolean '是否不允许增加
Private mfrmParent As Object

Private mblnChangeSel As Boolean

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Private Declare Function GetWindowRect Lib "user32" (ByVal Hwnd As Long, lpRect As RECT) As Long

'-----------------------------------------------------------------------------------------------------
Public Function ShowMe(ByVal frmParent As Object, ByVal lngHwnd As Long, ByVal intType As Integer, ByVal int期效 As Integer, ByVal int服务对象 As Integer, _
            Optional ByVal blnNewLIS As Boolean, Optional ByVal blnNew As Boolean, Optional ByVal lng项目id As Long, Optional ByRef strExtData As String) As Boolean
'参数:
'     frmParent         父窗体
'     lngHwnd           用于定位的控件句柄,即调用该窗体的控件
'     intType           0-检查组合,1-手术输入,4-检验组合，5-输血或治疗类需要填写申请附项的
'     int期效           将要输入的医嘱期效 0-长嘱，1-临嘱
'     int服务对象       该医嘱要服务的病人性质 1-门诊（包括门诊病人，体检病人，外来病人等) 2-住院（只有住院病人）
'     blnNewLIS         判断检验组合是否使用新版LIS的检验组合模式
'     blnNew            判断是否是新开输入项目时进入，否则为点下箭头进入。 true-新开输入项目时进入， false-点下箭头进入（现在只针对检验，只在新版LIS中使用（blnNewLIS=true)）
'     lng项目id         主诊疗项目ID
'返回：
'     strExtData        附加定义数据 , 新增时一般为空
'                       检查 = "部位名1;方法名1,方法名2|部位名2;方法名1,方法名2|...<vbTab>0-常规/1-床旁/2-术中"
'                       手术="手术ID1,手术ID2,...;麻醉ID",其中可能没有附加手术和麻醉
'                       检验组合="项目ID1,项目ID2,...;检验标本" 如果是新版LIS的模式则是："项目ID1|指标1|指标2...,项目ID2|指标1|指标2...,...;检验标本"


    Set mfrmParent = frmParent
    mlngHwnd = lngHwnd
    mintType = intType
    mint期效 = int期效
    mint服务对象 = int服务对象
    mblnNewLIS = blnNewLIS
    mblnNew = blnNew
    mlng项目ID = lng项目id
    mstrExtData = strExtData
    mblnOK = False
    On Error Resume Next
    Me.Show 1, frmParent
    
    strExtData = mstrExtData
    ShowMe = mblnOK
End Function

Private Sub cbo标本_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If cbo标本.ListIndex <> -1 Then
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    Else
        lngIdx = Cbo.MatchIndex(cbo标本.Hwnd, KeyAscii)
        If lngIdx = -1 And cbo标本.ListCount > 0 Then lngIdx = 0
        cbo标本.ListIndex = lngIdx
    End If
End Sub

Private Sub cmd_Click()
'功能：打开项目选择器
    Dim rsTmp As ADODB.Recordset, i As Long
    Dim strSql As String, strSQLItem As String
    Dim vPoint As PointAPI, blnCancel As Boolean, str药品 As String
    Dim strSamples As String
    
    On Error GoTo errH
    
    If mintType = 1 Then
        '输入附加手术:这里不是单独应用,因此不限制
        strSQLItem = _
            " From 诊疗项目目录 A Where A.类别='F' And A.ID<>" & mlng项目ID & _
                " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL)" & _
                " And (A.服务对象 IN([1],3) Or [1]=3 And Nvl(A.服务对象,0)<>0) And Nvl(A.执行频率,0) IN(0,[2])"
        
        strSql = "Select Distinct 0 as 末级,ID,上级ID,编码,名称,NULL as 单位,NULL as 规模" & _
            " From 诊疗分类目录 Where 类型=5 And (撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " Start With ID In (Select 分类ID" & strSQLItem & ") Connect by Prior 上级ID=ID"
        strSql = strSql & " Union ALL" & _
            " Select Distinct 1 as 末级,A.ID,分类ID as 上级ID,A.编码,A.名称,A.计算单位 as 单位,A.操作类型 as 规模" & _
            strSQLItem & " Order By 编码"
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSql, 2, "手术", False, "", "", False, True, False, 0, 0, 0, blnCancel, False, False, _
            mint服务对象, IIF(mint期效 = 0, 2, 1))
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "未找到可用的手术项目，请先到诊疗项目管理中设置。", vbInformation, gstrSysName
            End If
            Exit Sub
        End If
        
        '检查重复输入
        i = vsExt.FindRow(CLng(rsTmp!ID))
        If i <> -1 And i <> vsExt.Row Then
            MsgBox "该附加手术已经在其它行录入。", vbInformation, gstrSysName
            Exit Sub
        End If
        
        Call Set手术输入(vsExt.Row, rsTmp)
    ElseIf mintType = 4 Then
        '检验项目
        With Me.cbo标本
            For i = 0 To .ListCount - 1
                strSamples = strSamples & ",'" & .List(i) & "'"
            Next
        End With
        If Len(strSamples) > 0 Then
            strSamples = Mid(strSamples, 2)
        Else
            strSamples = "''"
        End If
        
        strSQLItem = "From 诊疗项目目录 A,检验项目参考 C,检验报告项目 D " & _
            "Where A.ID=D.诊疗项目id(+) And D.报告项目ID=C.项目id(+)" & _
            " And A.类别='C' And Nvl(A.单独应用,0)=1" & _
            " And (A.服务对象 IN([1],3) Or [1]=3 And Nvl(A.服务对象,0)<>0)" & _
            " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL)" & _
            " And (C.标本类型 In (" & strSamples & ") Or C.标本类型 Is Null)"
        
        strSql = "Select Distinct 0 as 末级,ID,上级ID,编码,名称,' ' As 检验类型,' ' As 标本部位" & _
            " From 诊疗分类目录 Where 类型=5 And (撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " Start With ID In (Select A.分类ID " & strSQLItem & ") Connect by Prior 上级ID=ID"
        strSql = strSql & " Union ALL" & _
            " Select Distinct 1 as 末级,A.ID,分类ID as 上级ID,A.编码,A.名称,A.操作类型 as 检验类型,A.标本部位 " & strSQLItem & " Order By 编码"
        
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSql, 2, "检验项目", False, "", "", False, True, False, 0, 0, 0, blnCancel, False, False, mint服务对象)
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "未找到可用的检验项目，请先到诊疗项目管理中设置。", vbInformation, gstrSysName
            End If
            Exit Sub
        End If
        If rsTmp!检验类型 = "微生物" And vsExt.Rows > 2 Then
            If vsExt.RowData(2) <> 0 Or vsExt.Row > 1 Then '整个申请只能开一个微生物项目
                MsgBox "微生物项目只能单独申请！", vbInformation, gstrSysName
                Exit Sub
            End If
        End If
        
        '检查重复输入
        i = vsExt.FindRow(CLng(rsTmp!ID))
        If i <> -1 And i <> vsExt.Row Then
            MsgBox "该检验项目已经录入！", vbInformation, gstrSysName
            Exit Sub
        End If
        
        '检查检验类型是否相同
        For i = 1 To vsExt.Rows - 1
            If vsExt.RowData(i) <> 0 And i <> vsExt.Row Then
                If Not (vsExt.TextMatrix(i, 1) = Nvl(rsTmp!检验类型) _
                    Or vsExt.TextMatrix(i, 1) = "" Or Nvl(rsTmp!检验类型) = "") Then
                    MsgBox "请输入相同检验类型的项目，已输入项目的检验类型为""" & vsExt.TextMatrix(i, 1) & """。", vbInformation, gstrSysName
                    Exit Sub
                End If
            End If
        Next
        
        '重新初始标本
        If Not InitCombox(rsTmp!ID, Nvl(rsTmp!标本部位)) Then Exit Sub
        
        Call Set检验项目(vsExt.Row, rsTmp)
        If rsTmp("检验类型") = "微生物" Then
            mblnNotAddNew = True
            vsExt.Rows = 2
        Else
            mblnNotAddNew = False
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdData_Click()
'功能：打开项目选择器
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnCancel As Boolean
    Dim strSQLItem As String
    
    If mintType = 1 Then
        '输入麻醉项目:这里不是单独应用,因此不限制
        strSQLItem = " From 诊疗项目目录 A Where A.类别='G'" & _
                " And (A.服务对象 IN([2],3) Or [2]=3 And Nvl(A.服务对象,0)<>0) And A.ID<>[1]" & _
                " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL)"

        strSql = "Select Distinct 0 as 末级,ID,上级ID,编码,名称,NULL as 单位,NULL as 麻醉类型" & _
            " From 诊疗分类目录 Where 类型=5 And (撤档时间 Is Null Or 撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " Start With ID In (Select 分类ID" & strSQLItem & ") Connect by Prior 上级ID=ID"
        strSql = strSql & " Union ALL" & _
            " Select Distinct 1 as 末级,A.ID,分类ID as 上级ID,A.编码,A.名称,A.计算单位 as 单位,A.操作类型 as 麻醉类型" & _
            strSQLItem & " Order By 编码"
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSql, 2, "麻醉项目", False, "", "", False, True, False, 0, 0, 0, blnCancel, False, False, _
            mlng项目ID, mint服务对象)
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "未找到匹配项目！", vbInformation, gstrSysName
            End If
            txtData.SetFocus: Exit Sub
        End If
        txtData.Tag = rsTmp!ID
        txtData.Text = "[" & rsTmp!编码 & "]" & rsTmp!名称
        cmdData.Tag = txtData.Text
        
        txtData.SetFocus
    ElseIf mintType = 4 Then
        '输入标本
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdMethodOK_Click()
    Call vsMethod_KeyPress(vbKeyReturn)
End Sub

Private Sub cmdOK_Click()
    Dim blnSkip As Boolean
    Dim strMsg As String, strTmp As String
    Dim strSql As String, i As Long, j As Long
    Dim rsTmp As ADODB.Recordset
    
    
    Dim lngBegin As Long, lngEnd As Long
    Dim strData As String
    
    If mintType = 0 Then '检查部位组合
        '收集部位及方法的情况
        With vsExt
            For i = .FixedRows To .Rows - 1
                If .Cell(flexcpData, i, 1) = 1 Then
                    If .TextMatrix(i, 2) = "" Then
                        .Row = i: .ShowCell .Row, .Col
                        MsgBox "没有为检查部位""" & .TextMatrix(i, 1) & """确定检查方法。", vbInformation, gstrSysName
                        vsExt.SetFocus: Exit Sub
                    End If
                    
                    strTmp = strTmp & "|" & .TextMatrix(i, 1) & ";" & .TextMatrix(i, 2)
                End If
            Next
            If strTmp = "" And vsExt.Editable <> flexEDNone Then
                MsgBox "请至少选择一个检查部位。", vbInformation, gstrSysName
                vsExt.SetFocus: Exit Sub
            End If
            strTmp = Mid(strTmp, 2) & vbTab & 0
        End With
    ElseIf mintType = 1 Or mintType = 4 Then '附加手术及麻醉项目；检验项目及标本
        If mintType = 1 Or mintType = 4 And mblnNewLIS = False Then
            For i = 1 To vsExt.Rows - 1
                If vsExt.RowData(i) <> 0 Then
                    strTmp = strTmp & "," & vsExt.RowData(i)
                End If
            Next
        ElseIf mintType = 4 And mblnNewLIS Then
            For i = 1 To vsExt.Rows - 1
                If vsExt.RowData(i) <> 0 And (Val(vsExt.Cell(flexcpChecked, i, 0)) = 1 Or Val(vsExt.TextMatrix(i, 3)) = 0) Then
                    strTmp = strTmp & IIF(Val(vsExt.TextMatrix(i, 3)) = 1, "|", ",") & vsExt.RowData(i)
                End If
            Next
        End If
        strTmp = Mid(strTmp, 2)
        If strTmp = "" And mintType = 4 Then
            MsgBox "至少要选择一个检验项目。", vbInformation, gstrSysName
            vsExt.SetFocus: Exit Sub
        End If
        strTmp = strTmp & ";" & IIF(mintType = 4, Me.cbo标本.Text, IIF(Val(txtData.Tag) = 0, "", Val(txtData.Tag)))
    End If
    
    mstrExtData = strTmp
    mblnOK = True
    Unload Me
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    
    If KeyCode = vbKeyEscape Then
        If fraMethod.Visible Then
            fraMethod.Visible = False
            vsExt.SetFocus
        Else
            Call cmdCancel_Click
        End If
    ElseIf KeyCode = vbKeyF2 Then
        If cmdOK.Enabled And cmdOK.Visible Then Call cmdOK_Click
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyA Then 'CTRL+A
        If mintType = 0 Then
            vsExt.Cell(flexcpData, vsExt.FixedRows, 1, vsExt.Rows - 1, 1) = 1
            Set vsExt.Cell(flexcpPicture, vsExt.FixedRows, 1, vsExt.Rows - 1, 1) = img16.ListImages("c1").Picture
        End If
    ElseIf Shift = vbCtrlMask And KeyCode = vbKeyR Then 'CTRL+R
        If mintType = 0 Then
            vsExt.Cell(flexcpData, vsExt.FixedRows, 1, vsExt.Rows - 1, 1) = 0
            Set vsExt.Cell(flexcpPicture, vsExt.FixedRows, 1, vsExt.Rows - 1, 1) = img16.ListImages("c0").Picture
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    '不允许输入分隔符及单引号
    If InStr(",;|'", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub Form_Resize()
    Dim lngMinRows As Long
    Dim lngRows As Long, i As Long
    
    On Error Resume Next
    
    fraBorder(0).Left = 0
    fraBorder(0).Top = 0
    fraBorder(0).Width = Me.ScaleWidth
    fraBorder(1).Top = fraBorder(0).Height
    fraBorder(1).Left = Me.ScaleWidth - fraBorder(1).Width
    fraBorder(1).Height = Me.ScaleHeight - fraBorder(0).Height * 2
    fraBorder(2).Left = 0
    fraBorder(2).Top = Me.ScaleHeight - fraBorder(2).Height
    fraBorder(2).Width = Me.ScaleWidth
    fraBorder(3).Top = fraBorder(0).Height
    fraBorder(3).Left = 0
    fraBorder(3).Height = Me.ScaleHeight - fraBorder(0).Height * 2
    
    vsExt.Left = fraBorder(3).Width
    vsExt.Top = fraBorder(0).Height + fraBorder(0).Height
    vsExt.Width = Me.ScaleWidth - fraBorder(3).Width * 2
    
    vsExt.Height = Me.ScaleHeight - fraBorder(2).Height * 2 - (cbo标本.Height + 200)
    
    cbo标本.Top = Me.ScaleHeight - fraBorder(2).Height - ((Me.ScaleHeight - fraBorder(0).Height * 2 - vsExt.Height) - cbo标本.Height) / 2 - cbo标本.Height
    
    txtData.Top = cbo标本.Top
    lblData.Top = cbo标本.Top + (cbo标本.Height - lblData.Height) / 2
    cmdOK.Top = cbo标本.Top + (cbo标本.Height - cmdOK.Height) / 2
    cmdCancel.Top = cmdOK.Top
    
    lblData.Left = 200
    cbo标本.Left = lblData.Left + lblData.Width + fraBorder(3).Width
    txtData.Left = cbo标本.Left
    cmdCancel.Left = Me.ScaleWidth - cmdCancel.Width - cmdCancel.Height
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - fraBorder(1).Width * 3
    
            
    cbo标本.Width = cmdOK.Left - cbo标本.Left - 200
    txtData.Width = cbo标本.Width
    cmdData.Top = txtData.Top + 30
    cmdData.Left = txtData.Left + txtData.Width - cmdData.Width - 45
    
    Me.Refresh
End Sub

Private Sub Form_Load()
    Dim blnMulti As Boolean, vRect As RECT
    Dim str方法 As String, i As Long
    
    Me.Height = 2325
    
    '边框设置
    For i = 0 To fraBorder.UBound
        fraBorder(i).BackColor = vbButtonFace
    Next
    Set lin(0).Container = fraBorder(0): Set lin(1).Container = fraBorder(0)
    Set lin(2).Container = fraBorder(1): Set lin(3).Container = fraBorder(1)
    Set lin(4).Container = fraBorder(2): Set lin(5).Container = fraBorder(2)
    Set lin(6).Container = fraBorder(3): Set lin(7).Container = fraBorder(3)
    lin(0).X1 = 0: lin(0).Y1 = 0: lin(0).X2 = Screen.Width: lin(0).Y2 = lin(0).Y1: lin(0).BorderColor = &H8000000F
    lin(1).X1 = 0: lin(1).Y1 = Screen.TwipsPerPixelY: lin(1).X2 = Screen.Width: lin(1).Y2 = lin(1).Y1: lin(1).BorderColor = &H8000000E
    lin(2).X1 = fraBorder(1).Width - Screen.TwipsPerPixelX: lin(2).Y1 = 0: lin(2).X2 = lin(2).X1: lin(2).Y2 = Screen.Height: lin(2).BorderColor = &H80000011
    lin(3).X1 = fraBorder(1).Width - Screen.TwipsPerPixelX * 2: lin(3).Y1 = 0: lin(3).X2 = lin(3).X1: lin(3).Y2 = Screen.Height: lin(3).BorderColor = &H80000010
    lin(4).X1 = 0: lin(4).Y1 = fraBorder(2).Height - Screen.TwipsPerPixelY: lin(4).X2 = Screen.Width: lin(4).Y2 = lin(4).Y1: lin(4).BorderColor = &H80000011
    lin(5).X1 = 0: lin(5).Y1 = fraBorder(2).Height - Screen.TwipsPerPixelY * 2: lin(5).X2 = Screen.Width: lin(5).Y2 = lin(5).Y1: lin(5).BorderColor = &H80000010
    lin(6).X1 = 0: lin(6).Y1 = 0: lin(6).X2 = lin(6).X1: lin(6).Y2 = Screen.Height: lin(6).BorderColor = &H8000000F
    lin(7).X1 = Screen.TwipsPerPixelX: lin(7).Y1 = 0: lin(7).X2 = lin(7).X1: lin(7).Y2 = Screen.Height: lin(7).BorderColor = &H8000000E
    
 
    If mint服务对象 = 0 Then mint服务对象 = 3
    mblnOK = False
    mblnNotAddNew = False
                
    '初始化表格样式
    If mintType = 0 Then
        If Not Init检查组合 Then Unload Me: Exit Sub
    ElseIf mintType = 1 Then
        lblData.Visible = True
        txtData.Visible = True
        cmdData.Visible = True
        lblData.Caption = "麻醉"
        If Not Init手术项目 Then Unload Me: Exit Sub
    ElseIf mintType = 4 Then
        lblData.Visible = True
        lblData.Caption = "标本"
        With cbo标本
            .Left = txtData.Left: .Top = txtData.Top: .Width = txtData.Width
            .Visible = True
        End With
        If Not Init检验组合 Then Unload Me: Exit Sub
        If Not InitCombox(DefaultValue:=Me.txtData) Then Unload Me: Exit Sub
    End If
    
    '其他处理
    If mintType = 0 Then
        If vsExt.Rows = vsExt.FixedRows + 1 Then
            If vsExt.Editable = flexEDNone Then
                '没有设置部位时，则自动确认
                Call cmdOK_Click: Exit Sub
            ElseIf vsExt.TextMatrix(vsExt.FixedRows, 1) <> "" Then
                '只有一个部位，且部位只有一个方法可选时，自动确认
                If vsExt.TextMatrix(vsExt.FixedRows, 1) <> "" Then
                    '只有一个部位，自动选中该部位
                    vsExt.Cell(flexcpData, vsExt.FixedRows, 1) = 1
                    Set vsExt.Cell(flexcpPicture, vsExt.FixedRows, 1) = img16.ListImages("c1").Picture
                    '如果没有默认方法，只有一个方法也选中
                    str方法 = GetOnlyOneMethod(vsExt.Cell(flexcpData, vsExt.FixedRows, 2))
                    If vsExt.TextMatrix(vsExt.FixedRows, 2) = "" And str方法 <> "" Then
                        vsExt.TextMatrix(vsExt.FixedRows, 2) = str方法
                    End If
                    If vsExt.TextMatrix(vsExt.FixedRows, 2) <> "" Then vsExt.TabStop = False
                    
                    '只有一个方法可选时，如果不需要输入申请附项，则界面也不弹出
                    If vsExt.TextMatrix(vsExt.FixedRows, 2) <> "" And str方法 <> "" Then
                        Call cmdOK_Click: Exit Sub
                    End If
                End If
            End If
        End If
    ElseIf mintType = 4 Then
        '检验输入的特殊处理
        blnMulti = Val(zlDatabase.GetPara(84, glngSys)) = 1 '是否允许一条医嘱申请多个检验项目
        If Len(Trim(mstrExtData)) > 0 Then
            If Len(Trim(Split(mstrExtData, ";")(0))) > 0 And Not blnMulti Then
                vsExt.Enabled = False
                '如果只有一个标本则不显示本窗口
                If cbo标本.ListCount < 2 Then cmdOK_Click: Exit Sub
            End If
        End If
    End If
    
    '恢复个性化
    Call RestoreWinState(Me, App.ProductName, mintType)
    
    
    '窗体定位
    GetWindowRect mlngHwnd, vRect
    Me.Left = (vRect.Left - 1) * Screen.TwipsPerPixelX
    Me.Top = (vRect.Top - 1) * Screen.TwipsPerPixelY - Me.Height
    Call Form_Resize
    
End Sub

Private Function Init手术项目() As Boolean
'功能：初始化手术表格格式及数据
'参数：mstrExtData=包含附加手术及麻醉项目的信息,其中可能没有附加手术；为空时表示新输入手术项目
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, lng麻醉ID As Long
    Dim arr手术IDs As Variant, str手术IDs As String
    Dim i As Long, j As Long
    
    On Error GoTo errH
    
    strSql = mstrExtData
    If strSql = "" Then strSql = ";"
    str手术IDs = CStr(Split(strSql, ";")(0))
    lng麻醉ID = Val(Split(strSql, ";")(1))
    
    '附加手术
    If str手术IDs <> "" Then
        strSql = "Select /*+ Rule*/ A.ID,A.编码,A.名称,A.操作类型" & _
            " From 诊疗项目目录 A" & _
            " Where A.类别='F' And A.ID IN(Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))" & _
            " Order by A.编码"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, str手术IDs)
        i = rsTmp.RecordCount
    End If
        
    vsExt.Clear
    vsExt.Rows = IIF(i = 0, 2, i + 1)
    vsExt.Cols = 2
    vsExt.FixedRows = 1: vsExt.FixedCols = 0
    vsExt.TextMatrix(0, 0) = "附加手术"
    vsExt.TextMatrix(0, 1) = "规模"
    vsExt.ColWidth(0) = 3200: vsExt.ColWidth(1) = 800
    vsExt.FixedAlignment(0) = 4: vsExt.FixedAlignment(1) = 4
    vsExt.ColAlignment(0) = 1: vsExt.ColAlignment(1) = 1
    vsExt.Editable = flexEDKbdMouse
    
    If str手术IDs <> "" And i <> 0 Then
        arr手术IDs = Split(str手术IDs, ",") '按照原有输入顺序
        For i = 0 To UBound(arr手术IDs)
            rsTmp.Filter = "ID=" & CStr(arr手术IDs(i))
            If Not rsTmp.EOF Then
                j = j + 1
                vsExt.RowData(j) = CLng(rsTmp!ID)
                vsExt.TextMatrix(j, 0) = "[" & rsTmp!编码 & "]" & rsTmp!名称
                vsExt.Cell(flexcpData, j, 0) = vsExt.TextMatrix(j, 0) '用于恢复显示
                vsExt.TextMatrix(j, 1) = Nvl(rsTmp!操作类型, 0)
            End If
        Next
    End If
    
    '麻醉项目
    If lng麻醉ID <> 0 Then
        strSql = "Select A.ID,A.编码,A.名称,操作类型 From 诊疗项目目录 A Where A.类别='G' And A.ID=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng麻醉ID)
        If rsTmp.Filter <> 0 Then rsTmp.Filter = 0
        If Not rsTmp.EOF Then
            txtData.Tag = rsTmp!ID
            txtData.Text = "[" & rsTmp!编码 & "]" & rsTmp!名称
            cmdData.Tag = txtData.Text '用于恢复显示
        End If
    End If
    
    vsExt.Row = 1: vsExt.Col = 0
    Init手术项目 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Init检查组合() As Boolean
'功能：初始化检查部位表格格式及数据
'参数：mstrExtData=包含检查部位的信息,为空时表示新输入检查组合项目
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, lngIdx As Long, i As Integer
    Dim str类型 As String, str名称 As String
    Dim arrData As Variant, strNoneRegion As String
    Dim blnNone As Boolean
    
    On Error GoTo errH
    
    '读取检查项目基本信息
    strSql = "Select 名称,操作类型,执行标记 From 诊疗项目目录 Where ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng项目ID)
    str类型 = rsTmp!操作类型
    str名称 = rsTmp!名称
    
    '读取检查部位信息
    strSql = "Select B.分组,A.部位,A.方法,A.默认,B.备注,B.方法 as 检查方法 From 诊疗项目部位 A,诊疗检查部位 B" & _
        " Where A.类型=B.类型 And A.部位=B.名称 And A.项目ID=[1] And A.类型=[2] Order by B.分组,B.编码"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng项目ID, str类型)
    blnNone = rsTmp.EOF
'    If rsTmp.EOF Then
'        '如果该检查项目还没有设置检查部位,则以所有的供选择
'        strSQL = "Select 分组,名称 as 部位,Null as 方法,Null as 默认,备注,方法 as 检查方法 From 诊疗检查部位 Where 类型=[1] Order by 分组,编码"
'        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str类型)
'        If rsTmp.EOF Then
'            MsgBox "该项目的检查类型""" & str类型 & """下面没有设置任何检查部位，请先到检查部位管理中进行设置。", vbInformation, gstrSysName
'            Exit Function
'        End If
'    End If
    With vsExt
        '显示基准的部位及默认方法
        If blnNone Then
            .HighLight = flexHighlightNever
            .Editable = flexEDNone
            .TabStop = False
        Else
            .HighLight = flexHighlightAlways
            .Editable = flexEDKbdMouse
        End If
        .WordWrap = True
        .FocusRect = flexFocusNone
        .BackColorSel = &HFFCC99
        .ForeColorSel = &H0&
        .FixedRows = 1: .FixedCols = 0
        .Rows = .FixedRows + 1: .Cols = 4
        .MergeCellsFixed = flexMergeFree: .MergeRow(0) = True
        .MergeCells = flexMergeFree: .MergeCol(0) = True
        If str类型 = "病理" Then
            .TextMatrix(0, 0) = "标本名称"
            .TextMatrix(0, 1) = "标本名称"
            .TextMatrix(0, 2) = "材料类别"
        Else
            .TextMatrix(0, 0) = "检查部位"
            .TextMatrix(0, 1) = "检查部位"
            .TextMatrix(0, 2) = "检查方法"
        End If
        
        .TextMatrix(0, 3) = "备注"
        .RowHeight(0) = 300
        .ColComboList(2) = "..."
        For i = 0 To .Cols - 1
            .FixedAlignment(i) = 4: .ColAlignment(i) = 1
        Next
        Do While Not rsTmp.EOF
            If .TextMatrix(.Rows - 1, 1) <> rsTmp!部位 Then
                If .TextMatrix(.Rows - 1, 1) <> "" Then
                    .Rows = .Rows + 1
                End If
                .TextMatrix(.Rows - 1, 0) = zlCommFun.GetNeedName("" & rsTmp!分组)
                .TextMatrix(.Rows - 1, 1) = rsTmp!部位
                Set .Cell(flexcpPicture, .Rows - 1, 1) = img16.ListImages("c0").Picture
                .Cell(flexcpData, .Rows - 1, 2) = CStr(Nvl(rsTmp!检查方法)) '供方法选择器使用
                .TextMatrix(.Rows - 1, 3) = Nvl(rsTmp!备注)
            End If
            If Nvl(rsTmp!默认, 0) = 1 Then '以"方法名1,方法名2,..."的方式显示部位检查方法
                .TextMatrix(.Rows - 1, 2) = .TextMatrix(.Rows - 1, 2) & "," & Nvl(rsTmp!方法)
                If Left(.TextMatrix(.Rows - 1, 2), 1) = "," Then
                    .TextMatrix(.Rows - 1, 2) = Mid(.TextMatrix(.Rows - 1, 2), 2)
                End If
            End If
            rsTmp.MoveNext
        Loop
        
        '修改时套入已有的内容
        '  如果为空，也可能是以前的单部位检查项目，这时要以新增的方式重新选择部位
        '  或者对于以前的单部位项目，强行传入以前的部位(没有方法)，现还可能有同名部位
        If mstrExtData <> "" Then
            arrData = Split(Split(mstrExtData, vbTab)(0), "|")
            For i = 0 To UBound(arrData)
                lngIdx = .FindRow(CStr(Split(arrData(i), ";")(0)), 1, 1, , True)
                If lngIdx <> -1 Then
                    '该部位的方法:可能以前的数据只有部位没有方法
                    If UBound(Split(arrData(i), ";")) >= 1 Then
                        .TextMatrix(lngIdx, 2) = Split(arrData(i), ";")(1)
                    Else
                        .TextMatrix(lngIdx, 2) = ""
                    End If
                    .Cell(flexcpData, lngIdx, 1) = 1 '表明该部位已选择
                    Set .Cell(flexcpPicture, lngIdx, 1) = img16.ListImages("c1").Picture
                Else
                    '该部位设置已不存在
                    strNoneRegion = strNoneRegion & "," & Split(arrData(i), ";")(0)
                End If
            Next
        End If
        
        .Row = 1: .Col = 1
        .ShowCell .Row, .Col
        
        '确定表格尺寸
        .AutoSize 0, .Cols - 1
        If .ColWidth(0) < 500 Then .ColWidth(0) = 500
        If .ColWidth(0) > 850 Then .ColWidth(0) = 850
        If .ColWidth(1) < 800 Then .ColWidth(1) = 800
        If .ColWidth(1) > 1600 Then .ColWidth(1) = 1600
        If .ColWidth(2) < 2500 Then .ColWidth(2) = 2500
        If .ColWidth(2) > 3500 Then .ColWidth(2) = 3500
        If .ColWidth(3) < 800 Then .ColWidth(3) = 800
        If .ColWidth(3) > 2000 Then .ColWidth(3) = 2000
        
        lngIdx = 0
        For i = 0 To .Cols - 1
            lngIdx = lngIdx + .ColWidth(i) + 15
        Next
        Me.Width = lngIdx + 90
        
        .Height = (.Rows - 1) * (.RowHeightMin + 15) + .RowHeight(0) + 60
        If Not blnNone Then
            If .Height < 1590 Then .Height = 1590 '最少5行部位
            If .Height > 2865 + 50 Then .Height = 2865 + 50 '最多10行部位
        End If
    End With
    
    Me.Height = (vsExt.Height + 90) + cmdOK.Height + (cmdOK.Height * 0.65)
    
    '已不存在的部位提示
    If strNoneRegion <> "" Then
        If str类型 = "病理" Then
            MsgBox "以下病理标本在项目设置中已不存在：" & vbCrLf & Mid(strNoneRegion, 2), vbInformation, gstrSysName
        Else
            MsgBox "以下检查部位在项目设置中已不存在：" & vbCrLf & Mid(strNoneRegion, 2), vbInformation, gstrSysName
        End If
    End If
    
    Init检查组合 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Init检验组合() As Boolean
'功能：初始化检验项目
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, blnLis As Boolean
    Dim arrItems As Variant, strItems As String
    Dim i As Long, j As Long
    Dim strLIS As String
    Dim strTmp As String
    Dim colTmp As New Collection
    Dim strItemTmp As String
    Dim lng父ID As Long
    Dim Y As Long
    
    On Error GoTo errH
    
    strSql = mstrExtData
    If strSql = "" Then strSql = IIF(mlng项目ID <> 0, mlng项目ID, "") & ";"
    strItems = CStr(Split(strSql, ";")(0))
    Me.txtData.Text = Split(strSql, ";")(1)
    cmdData.Tag = txtData.Text
    
    If strItems <> "" Then
        '判断是否是新版LIS模式的组合项目
        If Not gobjLIS Is Nothing Then
            blnLis = gobjLIS.CheckLisSate
        End If
        If mblnNewLIS And blnLis Then
            strLIS = " Union All" & vbNewLine & _
                    "       Select e.Id, e.编码, e.名称, e.操作类型, 检验组合项目.编码 As 序号,检验组合项目.id as 父ID " & vbNewLine & _
                    "       From 检验组合项目, 检验报告项目 C, 检验报告项目 D, 诊疗项目目录 E" & vbNewLine & _
                    "       Where 检验组合项目.Id = c.诊疗项目id And c.报告项目id = d.报告项目id And d.诊疗项目id = e.Id And e.组合项目 <> 1 And 检验组合项目.Id <> e.Id"
            '分解子项
            For i = 0 To UBound(Split(strItems, ","))
                strTmp = Split(strItems, ",")(i)
                If InStr(strTmp, "|") > 0 Then
                    colTmp.Add Mid(strTmp, InStr(strTmp, "|") + 1), "_" & Mid(strTmp, 1, InStr(strTmp, "|") - 1)
                    strItemTmp = strItemTmp & "," & Mid(strTmp, 1, InStr(strTmp, "|") - 1)
                Else
                    strItemTmp = strItemTmp & "," & strTmp
                End If
            Next
            strItems = Mid(strItemTmp, 2)
            Me.Height = Me.Height + 1200
            vsExt.Height = vsExt.Height + 1200
        End If
        strSql = "Select * From (With 检验组合项目 As (Select /*+ Rule*/ A.ID,A.编码,A.名称,A.操作类型, a.编码 As 序号,null as 父ID  From 诊疗项目目录 A " & _
            " Where A.类别='C' And Nvl(A.单独应用,0)=1" & _
            " And (A.服务对象 IN([2],3) Or [2]=3 And Nvl(A.服务对象,0)<>0)" & _
            " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL)" & _
            " And A.ID In(Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist))))" & _
            " Select * from 检验组合项目" & _
            strLIS & _
            ") Order by 序号,编码"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strItems, mint服务对象)
    End If
        
    vsExt.Clear
    If strItems <> "" Then
        vsExt.Rows = IIF(rsTmp.RecordCount = 0, 2, rsTmp.RecordCount + 1)
    Else
        vsExt.Rows = 2
    End If
    vsExt.Cols = 4
    vsExt.FixedRows = 1: vsExt.FixedCols = 0
    vsExt.TextMatrix(0, 2) = "检验项目"
    If mblnNewLIS Then
        vsExt.ColWidth(2) = 3700
        vsExt.ColWidth(0) = 300
    Else
        vsExt.ColWidth(2) = 4000
        vsExt.ColHidden(0) = True
    End If
    vsExt.ColHidden(1) = True
    vsExt.ColHidden(3) = True
    vsExt.FixedAlignment(2) = 4
    vsExt.ColAlignment(2) = 1
    vsExt.Editable = flexEDKbdMouse
    
    If Not rsTmp.EOF Then
        arrItems = Split(strItems, ",") '按照原有输入顺序
        For i = 0 To UBound(arrItems)
            rsTmp.Filter = "ID=" & arrItems(i)
            If Not rsTmp.EOF Then
                Y = vsExt.FindRow(CLng(rsTmp!ID))
                '重复的指标不加入
                If Y = -1 Then
                    j = j + 1
                    vsExt.RowData(j) = CLng(rsTmp!ID)
                    '主项默认勾选，且不能取消
                    vsExt.TextMatrix(j, 0) = " "
                    vsExt.Cell(flexcpBackColor, j, 0) = &H8000000F
                    vsExt.TextMatrix(j, 2) = "[" & rsTmp!编码 & "]" & rsTmp!名称
                    vsExt.Cell(flexcpData, j, 2) = vsExt.TextMatrix(j, 2) '用于恢复显示
                    vsExt.TextMatrix(j, 1) = Nvl(rsTmp!操作类型)
                    vsExt.TextMatrix(j, 3) = 0   '父项
'                    If Nvl(rsTmp!操作类型) = "微生物" Then mblnNotAddNew = True '微生物只能开一个检验项目
                End If
                If mblnNewLIS Then
                    lng父ID = CLng(rsTmp!ID)
                    rsTmp.Filter = "父ID=" & CLng(rsTmp!ID)
                    Do While Not rsTmp.EOF
                        Y = vsExt.FindRow(CLng(rsTmp!ID))
                        '重复的指标不加入
                        If Y = -1 Then
                            j = j + 1
                            vsExt.RowData(j) = CLng(rsTmp!ID)
                            On Error Resume Next
                            strItemTmp = ""
                            strItemTmp = colTmp("_" & lng父ID)
                            On Error GoTo errH
                            If InStr("|" & strItemTmp & "|", "|" & CLng(rsTmp!ID) & "|") > 0 Then
                                vsExt.Cell(flexcpChecked, j, 0) = 1
                            ElseIf strItemTmp = "" And mblnNew Then  '第一次进入默认勾选
                                vsExt.Cell(flexcpChecked, j, 0) = 1
                            Else
                                vsExt.Cell(flexcpChecked, j, 0) = 2
                            End If
                            '子项缩进
                            vsExt.TextMatrix(j, 2) = "    [" & rsTmp!编码 & "]" & rsTmp!名称
                            vsExt.Cell(flexcpData, j, 2) = vsExt.TextMatrix(j, 2) '用于恢复显示
                            vsExt.TextMatrix(j, 1) = Nvl(rsTmp!操作类型)
                            If Nvl(rsTmp!操作类型) = "微生物" Then mblnNotAddNew = True '微生物只能开一个检验项目
                            vsExt.TextMatrix(j, 3) = 1    '子项
                        Else
                            '如果重复的指标勾选了前面的指标未勾选，则删除前面的指标加载后面的指标
                            On Error Resume Next
                            strItemTmp = ""
                            strItemTmp = colTmp("_" & lng父ID)
                            On Error GoTo errH
                            If vsExt.Cell(flexcpChecked, Y, 0) = 1 And InStr("|" & strItemTmp & "|", "|" & CLng(rsTmp!ID) & "|") > 0 Then
                                vsExt.RemoveItem Y
                                vsExt.AddItem ""
                                vsExt.RowData(j) = CLng(rsTmp!ID)
                                vsExt.Cell(flexcpChecked, j, 0) = 1
                                '子项缩进
                                vsExt.TextMatrix(j, 2) = "    [" & rsTmp!编码 & "]" & rsTmp!名称
                                vsExt.Cell(flexcpData, j, 2) = vsExt.TextMatrix(j, 2) '用于恢复显示
                                vsExt.TextMatrix(j, 1) = Nvl(rsTmp!操作类型)
                                If Nvl(rsTmp!操作类型) = "微生物" Then mblnNotAddNew = True '微生物只能开一个检验项目
                                vsExt.TextMatrix(j, 3) = 1    '子项
                            End If
                        End If
                        rsTmp.MoveNext
                    Loop
                End If
            End If
        Next
    End If
    If j > 0 Then vsExt.Rows = j + 1
    
    vsExt.Row = 1: vsExt.Col = 2
    Init检验组合 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function InitCombox(Optional ByVal strNewItemID As String = "", Optional ByVal DefaultValue As String = "") As Boolean
    Dim strSql As String
    Dim rsTmp As New ADODB.Recordset
    Dim i As Long, strTmp As String, lngItemCount As Long
    InitCombox = False
    
    On Error GoTo DBError
    strTmp = "": lngItemCount = 0
    For i = 1 To vsExt.Rows - 1
        If vsExt.RowData(i) <> 0 And (i <> vsExt.Row Or Len(strNewItemID) = 0) Then
            lngItemCount = lngItemCount + 1
            strTmp = strTmp & "," & vsExt.RowData(i)
        End If
    Next
    If Len(strNewItemID) > 0 Then
        lngItemCount = lngItemCount + 1
        strTmp = strTmp & "," & strNewItemID
    End If
    If Len(strTmp) > 0 Then strTmp = Mid(strTmp, 2)

    If lngItemCount = 0 Then
        strSql = "Select 名称 From 诊疗检验标本"
    Else
        strSql = "Select /*+ Rule*/ 标本类型,Sum(1) From (" & _
            "   Select Distinct A.ID,B.名称 As 标本类型" & _
            "   From 诊疗项目目录 A,诊疗检验标本 B,检验项目参考 C,检验报告项目 D" & _
            "   Where A.ID=D.诊疗项目ID(+) And D.报告项目ID=C.项目ID(+)" & _
            "        And (NVL(C.标本类型,'') Is Null Or NVL( C.标本类型,'')=B.名称)  And A.ID In (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist)))" & _
            " ) Group By 标本类型 Having Sum(1)=[2]"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strTmp, lngItemCount)
    If rsTmp.EOF Then
        MsgBox Switch(lngItemCount = 0, "未设置检验标本，请到字典管理工具中设置。", _
            lngItemCount = 1, "选取的检验项目未定义检验标本，请先到检验项目管理中设置", _
            lngItemCount > 1, "选取的检验项目的检验标本与其他项目的不一致，请先到检验项目管理中设置"), vbInformation, gstrSysName
        Exit Function
    End If
    
    With cbo标本
        strTmp = .Text
        
        .Clear
        Do While Not rsTmp.EOF
            .AddItem rsTmp(0)
            rsTmp.MoveNext
        Loop
        .ListIndex = 0
        On Error Resume Next
        If Len(DefaultValue) > 0 Then
            .Text = DefaultValue
        Else
            .Text = strTmp
        End If
    End With
    InitCombox = True
    Exit Function
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName, mintType)
    
    mlngHwnd = 0
    mintType = 0
    mint期效 = 0
    mint服务对象 = 0
    mblnNewLIS = False
    mblnNew = False
    mlng项目ID = 0
    Set mfrmParent = Nothing
End Sub

Private Sub fraBorder_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
    On Error Resume Next
    
    If Button = 1 Then
        If Index = 0 Then
            If Me.Height - Y < 2355 Or Me.Height - Y > 7200 Then Exit Sub
            Me.Top = Me.Top + Y
            Me.Height = Me.Height - Y
        ElseIf Index = 1 Then
            If Me.Width + x < 4140 Or Me.Width + x > 9600 Then Exit Sub
            Me.Width = Me.Width + x
        End If
        Call Form_Resize
    End If
End Sub

Private Sub txtData_GotFocus()
    zlControl.TxtSelAll txtData
End Sub

Private Sub txtData_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset, vRect As RECT
    Dim strSql As String, strLike As String
    Dim blnCancel As Boolean
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txtData.Text = "" Then
            If mintType = 1 Then '手术可以不输入麻醉项目
                Call zlCommFun.PressKey(vbKeyTab)
            End If
            Exit Sub
        ElseIf txtData.Text = cmdData.Tag Then
            Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        
        '优化
        strLike = gstrLike
        If Len(txtData.Text) < 2 Then strLike = ""
        
        If mintType = 1 Then
            '输入麻醉项目
            strSql = _
                " Select Distinct A.ID,A.编码,A.名称,A.计算单位 as 单位,A.操作类型 as 麻醉类型" & _
                " From 诊疗项目目录 A,诊疗项目别名 B" & _
                " Where A.ID=B.诊疗项目ID And A.类别='G' And (A.服务对象 IN([3],3) Or [3]=3 And Nvl(A.服务对象,0)<>0)" & _
                    " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL)" & _
                    " And (A.编码 Like [1] Or B.名称 Like [2] Or B.简码 Like [2]) And B.码类=[4]" & _
                " Order by A.编码"
            vRect = zlControl.GetControlRect(txtData.Hwnd)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSql, 0, "麻醉项目", False, "", "", False, False, True, vRect.Left, vRect.Top, txtData.Height, blnCancel, False, True, _
                UCase(txtData.Text) & "%", strLike & UCase(txtData.Text) & "%", mint服务对象, gbytCode + 1)
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "未找到匹配项目！", vbInformation, gstrSysName
                End If
                txtData.Text = cmdData.Tag
                zlControl.TxtSelAll txtData
                Exit Sub
            End If
            txtData.Tag = rsTmp!ID
            txtData.Text = "[" & rsTmp!编码 & "]" & rsTmp!名称
            cmdData.Tag = txtData.Text
            
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf mintType = 4 Then
            '检验标本
        End If
    ElseIf KeyAscii = Asc("*") Then
        KeyAscii = 0
        Call cmdData_Click
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txtData_Validate(Cancel As Boolean)
'功能：恢复显示原内容
    If txtData.Text <> cmdData.Tag Then
        txtData.Text = cmdData.Tag
    End If
End Sub

Private Sub vsExt_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
'功能:显示选择按钮,并保证当前单元格可见
     Dim strKey As String, lng药名ID As Long
     
    If mblnChangeSel = True Then Exit Sub
    '保证当前单元格可见
    If NewRow >= vsExt.FixedRows And NewRow <= vsExt.Rows - 1 Then
        If vsExt.LeftCol >= vsExt.FixedCols And vsExt.LeftCol <= vsExt.Cols - 1 Then
            Call vsExt.ShowCell(NewRow, vsExt.LeftCol)
        End If
    End If
    
    If mintType = 1 Or mintType = 4 Then
        '显示/隐藏手术选择按钮
        If NewCol = 0 And mintType = 1 Or NewCol = 2 And mintType = 4 Then
            cmd.Height = vsExt.CellHeight - 30
            cmd.Left = vsExt.CellLeft + vsExt.CellWidth - cmd.Width - 15
            cmd.Top = vsExt.CellTop + 15
            
            If mintType = 4 And mblnNewLIS Then
                If vsExt.TextMatrix(NewRow, 3) = "1" Then
                    cmd.Visible = False
                Else
                    cmd.Visible = True
                End If
            Else
                cmd.Visible = True
            End If
        Else
            cmd.Visible = False
        End If
        If cmd.Visible Then
            vsExt.FocusRect = flexFocusSolid
        Else
            vsExt.FocusRect = flexFocusLight
        End If
    End If
    
End Sub

Private Sub vsExt_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
'功能:限制某些列宽的范围
    If Row = -1 Then
        If mintType = 1 Or mintType = 4 Then
            Call vsExt_AfterRowColChange(-1, -1, vsExt.Row, vsExt.Col) '使按钮可见及调整按钮位置
        End If
    End If
End Sub

Private Sub vsExt_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If mintType = 0 Then
        If NewCol = 0 Or NewCol = 3 Then
            Cancel = True
            If NewRow <> OldRow Then vsExt.Row = NewRow
        End If
    End If
End Sub

Private Sub vsExt_BeforeScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long, Cancel As Boolean)
    If cmd.Visible Then cmd.Visible = False
    If fraMethod.Visible Then fraMethod.Visible = False
End Sub

Private Function GetOnlyOneMethod(ByVal strMethod As String) As String
'功能：根据部位的方法定义，如果只有一个方法可选，则返回该方法
    Dim strTmp As String
    
    If strMethod = "" Then Exit Function
    strTmp = strMethod
    
    strTmp = Replace(strTmp, vbTab, ";")
    strTmp = Replace(strTmp, ",", ";")
    strTmp = Replace(strTmp, ";;", ";")
    strTmp = "<spdel>" & strTmp & "<spdel>"
    strTmp = Replace(strTmp, "<spdel>;", "")
    strTmp = Replace(strTmp, ";<spdel>", "")
    strTmp = Replace(strTmp, "<spdel>", "")
    
    If InStr(strTmp, ";") = 0 Then GetOnlyOneMethod = Mid(strTmp, 2)        '去掉前首位造影标记
End Function

Private Sub vsExt_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim strMethod As String, i As Long, j As Long
    Dim arrMethod As Variant, arrSub As Variant
    
    strMethod = vsExt.Cell(flexcpData, Row, Col)
    If strMethod = "" Then
        MsgBox "该检查部位没有设置可供选择的检查方法。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    With vsMethod
        .Rows = 0
        arrMethod = Split(Replace(strMethod, vbTab, ";" & vbTab), ";")
        For i = 0 To UBound(arrMethod)
            arrSub = Split(arrMethod(i), ",")
            For j = 0 To UBound(arrSub)
                .Rows = .Rows + 1
                If j = 0 Then
                    If InStr(1, arrMethod(i), vbTab) > 0 Then
                        .MergeRow(.Rows - 1) = True
                        .RowData(.Rows - 1) = 2 '表明是共选项
                        
                        .Cell(flexcpText, .Rows - 1, 0, .Rows - 1, 1) = Mid(arrSub(j), 3) '第一位是造影剂标志
                        If InStr("," & vsExt.TextMatrix(vsExt.Row, 2) & ",", "," & Mid(arrSub(j), 3) & ",") > 0 Then
                            Set .Cell(flexcpPicture, .Rows - 1, 0, .Rows - 1, 1) = img16.ListImages("c1").Picture
                            .Cell(flexcpData, .Rows - 1, 0) = 1
                        Else
                            Set .Cell(flexcpPicture, .Rows - 1, 0, .Rows - 1, 1) = img16.ListImages("c0").Picture
                            .Cell(flexcpData, .Rows - 1, 0) = 0
                        End If
                    Else
                        '排斥项
                        .MergeRow(.Rows - 1) = True
                        .RowData(.Rows - 1) = 1 '表明是排斥项
                        
                        .Cell(flexcpText, .Rows - 1, 0, .Rows - 1, 1) = Mid(arrSub(j), 2) '第一位是造影剂标志
                        If InStr("," & vsExt.TextMatrix(vsExt.Row, 2) & ",", "," & Mid(arrSub(j), 2) & ",") > 0 Then
                            Set .Cell(flexcpPicture, .Rows - 1, 0, .Rows - 1, 1) = img16.ListImages("o1").Picture
                            .Cell(flexcpData, .Rows - 1, 0) = 1 '1为选中
                        Else
                            Set .Cell(flexcpPicture, .Rows - 1, 0, .Rows - 1, 1) = img16.ListImages("o0").Picture
                            .Cell(flexcpData, .Rows - 1, 0) = 0
                        End If
                    End If
                Else
                    '共选子项
                    .RowData(.Rows - 1) = 3 '表明是共选子项
                    
                    .Cell(flexcpText, .Rows - 1, 1) = Mid(arrSub(j), 2)
                    If InStr("," & vsExt.TextMatrix(vsExt.Row, 2) & ",", "," & Mid(arrSub(j), 2) & ",") > 0 Then
                        Set .Cell(flexcpPicture, .Rows - 1, 1) = img16.ListImages("c1").Picture
                        .Cell(flexcpData, .Rows - 1, 0) = 1
                    Else
                        Set .Cell(flexcpPicture, .Rows - 1, 1) = img16.ListImages("c0").Picture
                        .Cell(flexcpData, .Rows - 1, 0) = 0
                    End If
                End If
            Next
        Next
        
        .Row = 0: .Col = 1
        
        .Height = .Rows * (.RowHeightMin + 15) + 30
        If .Height > Me.ScaleHeight - 100 Then .Height = Me.ScaleHeight - 100
        If .Height < 3 * (.RowHeightMin + 15) + 30 Then .Height = 3 * (.RowHeightMin + 15) + 30
        
        .Width = (vsExt.Width - 30) - (vsExt.CellLeft + 15)
        .Left = vsExt.Left + vsExt.CellLeft + 15
        .Top = vsExt.Top + vsExt.CellTop + vsExt.CellHeight + 15
        
        If .Top + .Height > Me.ScaleHeight Then
            .Top = Me.ScaleHeight - .Height
        End If
        
        fraMethod.Top = .Top: .Top = 0
        fraMethod.Left = .Left: .Left = 0
        fraMethod.Width = .Width
        fraMethod.Height = .Height + cmdMethodOK.Height + 20
        cmdMethodOK.Top = .Height
        cmdMethodOK.Left = .Width - cmdMethodOK.Width - 20
        
        fraMethod.ZOrder
        fraMethod.Visible = True
        If fraMethod.Visible Then .SetFocus
    End With
End Sub

Private Sub vsExt_DblClick()
    If mintType = 0 Then
        If vsExt.Editable <> flexEDNone And vsExt.MouseCol = 1 And vsExt.MouseRow >= vsExt.FixedRows Then
            Call vsExt_KeyPress(vbKeySpace)
        End If
    End If
End Sub

Private Sub vsExt_GotFocus()
    If fraMethod.Visible Then fraMethod.Visible = False
    Call vsExt_AfterRowColChange(-1, -1, vsExt.Row, vsExt.Col) '使按钮可见
End Sub

Private Sub vsExt_KeyDown(KeyCode As Integer, Shift As Integer)
'功能：删除数据行
    Dim i As Long, j As Long, k As Long
    Dim intRow As Integer        '有效行
    Dim lngRow As Long
    Dim lngCol As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim strKey As String, lng药品ID As Long
    
   If KeyCode = vbKeyDelete Then
        If (mintType = 1 Or mintType = 4) And vsExt.RowData(vsExt.Row) <> 0 Then
            '如果是新版LIS组合项目模式，则不允许删除子项
            If mintType = 4 And mblnNewLIS Then
                If vsExt.TextMatrix(vsExt.Row, 3) = "1" Then Exit Sub
            End If
            If MsgBox("要删除当前行吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            
            '如果组合项目模式，则同时删除子项
            If mintType = 4 And mblnNewLIS Then
                lngBegin = vsExt.Row + 1
                For j = vsExt.Row + 1 To vsExt.Rows - 1
                    If vsExt.TextMatrix(j, 3) <> "1" Then Exit For
                    lngEnd = j
                Next
                For j = lngEnd To lngBegin Step -1
                    vsExt.RowData(j) = 0
                    For i = 0 To vsExt.Cols - 1
                        vsExt.TextMatrix(j, i) = ""
                        vsExt.Cell(flexcpData, j, i) = ""
                    Next
                    If Not (vsExt.Rows = vsExt.FixedRows + 1 And j = vsExt.FixedRows) Then
                        vsExt.RemoveItem j
                    End If
                Next
            End If
            vsExt.RowData(vsExt.Row) = 0
            For i = 0 To vsExt.Cols - 1
                vsExt.TextMatrix(vsExt.Row, i) = ""
                vsExt.Cell(flexcpData, vsExt.Row, i) = ""
            Next
            If Not (vsExt.Rows = vsExt.FixedRows + 1 And vsExt.Row = vsExt.FixedRows) Then
                vsExt.RemoveItem vsExt.Row
            End If
            
            '重新初始标本
            If mintType = 4 Then InitCombox
        End If
    End If
End Sub

Private Sub vsExt_LostFocus()
    If Not ActiveControl Is cmd Then cmd.Visible = False
End Sub

Private Sub vsExt_KeyPress(KeyAscii As Integer)
'功能：非编辑状态时，自动移动单元格
    If KeyAscii = 13 Then
        KeyAscii = 0
        '定位到下一应输入单元格
        If mintType = 0 Then
            If vsExt.Col <= 1 Then
                vsExt.Col = vsExt.Col + 1
            ElseIf vsExt.Col = 2 And vsExt.Row <= vsExt.Rows - 2 Then
                vsExt.Row = vsExt.Row + 1
                vsExt.Col = 1
            ElseIf vsExt.Col = 2 And vsExt.Row = vsExt.Rows - 1 Then
                Call zlCommFun.PressKey(vbKeyTab)
                Exit Sub
            End If
        ElseIf mintType = 1 Or mintType = 4 Then
            If vsExt.Row = vsExt.Rows - 1 Then
                If vsExt.RowData(vsExt.Row) = 0 Or mblnNotAddNew Then
                    Call zlCommFun.PressKey(vbKeyTab)
                    Exit Sub
                Else
                    vsExt.AddItem ""
                End If
            End If
            If vsExt.Row + 1 <= vsExt.Rows - 1 Then
                vsExt.Row = vsExt.Row + 1
                If mintType = 1 Then
                    vsExt.Col = 0
                Else
                    vsExt.Col = 2
                End If
            End If
        End If
    ElseIf KeyAscii = Asc("*") Then
        If mintType = 0 Then
            If vsExt.Col = 2 Then
                Call vsExt_CellButtonClick(vsExt.Row, vsExt.Col)
            End If
        ElseIf mintType = 1 Or mintType = 4 Then
            KeyAscii = 0
            If cmd.Visible Then cmd_Click
        End If
    ElseIf KeyAscii = vbKeySpace Then
        If mintType = 0 Then
            If vsExt.Editable <> flexEDNone Then
                If vsExt.Col = 1 Then
                    If vsExt.Cell(flexcpData, vsExt.Row, vsExt.Col) = 1 Then
                        vsExt.Cell(flexcpData, vsExt.Row, vsExt.Col) = 0
                        Set vsExt.Cell(flexcpPicture, vsExt.Row, vsExt.Col) = img16.ListImages("c0").Picture
                    Else
                        vsExt.Cell(flexcpData, vsExt.Row, vsExt.Col) = 1
                        Set vsExt.Cell(flexcpPicture, vsExt.Row, vsExt.Col) = img16.ListImages("c1").Picture
                        
                        '自动弹出方法选择器
                        vsExt.Col = 2
                        Call vsExt_CellButtonClick(vsExt.Row, vsExt.Col)
                    End If
                ElseIf vsExt.Col = 2 Then
                    Call vsExt_CellButtonClick(vsExt.Row, vsExt.Col)
                End If
            End If
        End If
    End If
End Sub

Private Sub vsExt_AfterEdit(ByVal Row As Long, ByVal Col As Long)
'功能：非回车确认完后编辑的处理(这里Text:=EditText,但ValidateEdit事件中还没有)
    Dim strKey As String, lng药名ID As Long, i As Long
    
    If Not mblnReturn Then
        If mintType = 1 Or mintType = 4 Then
            If Col = 0 And mintType = 1 Or Col = 2 And mintType = 4 Then
                vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                Call vsExt_AfterRowColChange(Row, Col, Row, Col) '重新使按钮可见
                            
                '重新初始标本
                If mintType = 4 Then InitCombox
            End If
        End If
    End If
End Sub


Private Sub vsExt_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
'功能：输入数据确认
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, strSamples As String
    Dim blnCancel As Boolean, i As Long
    Dim vPoint As PointAPI, strLike As String, str药品 As String
    Dim strKey As String, lng药名ID As Long
    
    If KeyAscii = 13 Then
        mblnReturn = True '标记是按回车确认编辑
        KeyAscii = 0
        
        '优化
        strLike = gstrLike
        If Len(vsExt.EditText) < 2 Then strLike = ""
        
        On Error GoTo errH
        
        If mintType = 1 Then
            '输入附加手术:这里不是单独应用,因此不限制
            strSql = _
                " Select Distinct A.ID,A.编码,A.名称,A.计算单位 as 单位,A.操作类型 as 规模" & _
                " From 诊疗项目目录 A,诊疗项目别名 B" & _
                " Where A.ID=B.诊疗项目ID And A.类别='F' And A.ID<>[3]" & IIF(strLike = "", "", " And Rownum<=100") & _
                    " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL)" & _
                    " And (A.编码 Like [1] Or B.名称 Like [2] Or B.简码 Like [2]) And B.码类=[4]" & _
                    " And (A.服务对象 IN([5],3) Or [5]=3 And Nvl(A.服务对象,0)<>0) And Nvl(A.执行频率,0) IN(0,[6])" & _
                " Order by A.编码"
            vPoint = zlControl.GetCoordPos(vsExt.Hwnd, vsExt.CellLeft, vsExt.CellTop)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSql, 0, "手术", False, "", "", False, False, True, vPoint.x, vPoint.Y, vsExt.CellHeight, blnCancel, False, True, _
                UCase(vsExt.EditText) & "%", strLike & UCase(vsExt.EditText) & "%", mlng项目ID, gbytCode + 1, mint服务对象, IIF(mint期效 = 0, 2, 1))
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "未找到匹配项目！", vbInformation, gstrSysName
                End If
                vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                Call vsExt_AfterRowColChange(Row, Col, Row, Col) '重新使按钮可见
                Exit Sub
            End If
            
            '检查重复输入
            i = vsExt.FindRow(CLng(rsTmp!ID))
            If i <> -1 And i <> Row Then
                MsgBox "该附加手术已经在其它行录入。", vbInformation, gstrSysName
                vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                Call vsExt_AfterRowColChange(Row, Col, Row, Col) '重新使按钮可见
                Exit Sub
            End If
            
            Call Set手术输入(Row, rsTmp)
        ElseIf mintType = 4 Then
            '检验项目
            With Me.cbo标本
                For i = 0 To .ListCount - 1
                    strSamples = strSamples & ",'" & .List(i) & "'"
                Next
            End With
            If Len(strSamples) > 0 Then
                strSamples = Mid(strSamples, 2)
            Else
                strSamples = "''"
            End If
            strSql = "Select A.ID,A.编码,A.名称,A.操作类型,A.标本部位" & _
                " From 诊疗项目目录 A,诊疗项目别名 C Where A.ID=C.诊疗项目ID" & _
                " And (A.编码 Like [1] Or C.名称 Like [2] Or C.简码 Like [2]) And C.码类=[3]" & _
                " And A.类别='C' And Nvl(A.单独应用,0)=1" & _
                " And (A.服务对象 IN([4],3) Or [4]=3 And Nvl(A.服务对象,0)<>0)" & _
                " And (A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD') Or A.撤档时间 IS NULL)"
            If strLike = "" Then
                '当可以利用简码索引时(单向匹配),如果有(+)连接,则需要Group By一下(奇怪)
                strSql = strSql & " Group by A.ID,A.编码,A.名称,A.操作类型,A.标本部位"
            End If
            
            strSql = "Select Distinct A.ID,A.编码,A.名称,A.操作类型 as 检验类型,A.标本部位" & _
                " From 检验项目参考 D,检验报告项目 E,(" & strSql & ") A" & _
                " Where A.ID=E.诊疗项目id(+) And E.报告项目ID=D.项目id(+)" & _
                " And (D.标本类型 In (" & strSamples & ") Or D.标本类型 Is Null)" & _
                " Order by A.编码"

            vPoint = zlControl.GetCoordPos(vsExt.Hwnd, vsExt.CellLeft, vsExt.CellTop)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSql, 0, "检验项目", False, "", "", False, False, True, vPoint.x, vPoint.Y, vsExt.CellHeight, blnCancel, False, True, _
                UCase(vsExt.EditText) & "%", strLike & UCase(vsExt.EditText) & "%", gbytCode + 1, mint服务对象)
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "未找到匹配项目！", vbInformation, gstrSysName
                End If
                vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                Call vsExt_AfterRowColChange(Row, Col, Row, Col) '重新使按钮可见
                Exit Sub
            End If
            If rsTmp!检验类型 = "微生物" And vsExt.Rows > 2 Then
                If vsExt.RowData(2) <> 0 Or vsExt.Row > 1 Then '整个申请只能开一个微生物项目
                    MsgBox "微生物项目只能单独申请！", vbInformation, gstrSysName
                    vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                    Call vsExt_AfterRowColChange(Row, Col, Row, Col) '重新使按钮可见
                    Exit Sub
                End If
            End If
            
            '检查重复输入
            i = vsExt.FindRow(CLng(rsTmp!ID))
            If i <> -1 And i <> Row Then
                MsgBox "该检验项目已经录入！", vbInformation, gstrSysName
                vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                Call vsExt_AfterRowColChange(Row, Col, Row, Col) '重新使按钮可见
                Exit Sub
            End If
            
            '检查检验类型是否相同
            For i = 1 To vsExt.Rows - 1
                If vsExt.RowData(i) <> 0 And i <> Row Then
                    If Not (vsExt.TextMatrix(i, 1) = Nvl(rsTmp!检验类型) _
                        Or vsExt.TextMatrix(i, 1) = "" Or Nvl(rsTmp!检验类型) = "") Then
                        MsgBox "请输入相同检验类型的项目，已输入项目的检验类型为""" & vsExt.TextMatrix(i, 1) & """。", vbInformation, gstrSysName
                        vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                        Call vsExt_AfterRowColChange(Row, Col, Row, Col) '重新使按钮可见
                        Exit Sub
                    End If
                End If
            Next
            
            '重新初始标本
            If Not InitCombox(rsTmp!ID, Nvl(rsTmp!标本部位)) Then
                vsExt.TextMatrix(Row, Col) = CStr(vsExt.Cell(flexcpData, Row, Col))
                Call vsExt_AfterRowColChange(Row, Col, Row, Col) '重新使按钮可见
                Exit Sub
            End If
            
            Call Set检验项目(Row, rsTmp)
            If rsTmp!检验类型 = "微生物" Then
                mblnNotAddNew = True
                vsExt.Rows = 2
            Else
                mblnNotAddNew = False
            End If
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Set手术输入(ByVal lngRow As Long, rsInput As ADODB.Recordset)
    '附加手术
    vsExt.EditText = "[" & rsInput!编码 & "]" & rsInput!名称 '对于输入直接匹配时有必要
    
    vsExt.RowData(lngRow) = CLng(rsInput!ID)
    vsExt.TextMatrix(lngRow, 0) = "[" & rsInput!编码 & "]" & rsInput!名称
    vsExt.Cell(flexcpData, lngRow, 0) = vsExt.TextMatrix(lngRow, 0)
    vsExt.TextMatrix(lngRow, 1) = Nvl(rsInput!规模)
    
    '下一输入行
    If vsExt.RowData(vsExt.Rows - 1) <> 0 And Not mblnNotAddNew Then vsExt.AddItem ""
    vsExt.Row = vsExt.Rows - 1: vsExt.Col = 0
End Sub

Private Sub Set检验项目(ByVal lngRow As Long, rsInput As ADODB.Recordset)
    Dim strSql As String, rsTmp As Recordset
    Dim i As Long, j As Long
    Dim lngBegin As Long, lngEnd As Long
    
    '检验项目
    '如果新LIS组合项目模式则先删除子项再路径
    '如果组合项目模式，则同时删除子项
    If mblnNewLIS Then
        lngBegin = lngRow + 1
        For j = lngRow + 1 To vsExt.Rows - 1
            If vsExt.TextMatrix(j, 3) <> "1" Then Exit For
            lngEnd = j
        Next
        For j = lngEnd To lngBegin Step -1
            vsExt.RowData(j) = 0
            For i = 0 To vsExt.Cols - 1
                vsExt.TextMatrix(j, i) = ""
                vsExt.Cell(flexcpData, j, i) = ""
            Next
            If Not (vsExt.Rows = vsExt.FixedRows + 1 And j = vsExt.FixedRows) Then
                vsExt.RemoveItem j
            End If
        Next
    End If
    vsExt.EditText = "[" & rsInput!编码 & "]" & rsInput!名称 '对于输入直接匹配时有必要
    
    vsExt.RowData(lngRow) = CLng(rsInput!ID)
    vsExt.TextMatrix(lngRow, 2) = "[" & rsInput!编码 & "]" & rsInput!名称
    vsExt.Cell(flexcpData, lngRow, 2) = vsExt.TextMatrix(lngRow, 2)
    vsExt.TextMatrix(lngRow, 1) = Nvl(rsInput!检验类型)
    vsExt.TextMatrix(lngRow, 0) = " "
    vsExt.Cell(flexcpBackColor, lngRow, 0) = &H8000000F
    vsExt.TextMatrix(lngRow, 3) = 0 '父项
    
    If mblnNewLIS Then
        strSql = "" & vbNewLine & _
            "       Select e.Id, e.编码, e.名称, e.操作类型, a.编码 As 序号, a.Id As 父id" & vbNewLine & _
            "       From 诊疗项目目录 a, 检验报告项目 C, 检验报告项目 D, 诊疗项目目录 E" & vbNewLine & _
            "       Where a.Id = c.诊疗项目id And c.报告项目id = d.报告项目id And d.诊疗项目id = e.Id And e.组合项目 <> 1 And a.Id <> e.Id and a.id=[1]" & vbNewLine & _
            "       Order By 序号, 编码"
        On Error GoTo errH
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, CLng(rsInput!ID))
        Do While Not rsTmp.EOF
            i = vsExt.FindRow(CLng(rsTmp!ID))
            '重复的指标不加入
            If i = -1 Then
                If vsExt.RowData(vsExt.Rows - 1) & "" <> "" Then vsExt.AddItem ""
                vsExt.RowData(vsExt.Rows - 1) = CLng(rsTmp!ID)
                vsExt.Cell(flexcpChecked, vsExt.Rows - 1, 0) = 1
                '子项缩进
                vsExt.TextMatrix(vsExt.Rows - 1, 2) = "    [" & rsTmp!编码 & "]" & rsTmp!名称
                vsExt.Cell(flexcpData, vsExt.Rows - 1, 2) = vsExt.TextMatrix(vsExt.Rows - 1, 2) '用于恢复显示
                vsExt.TextMatrix(vsExt.Rows - 1, 1) = Nvl(rsTmp!操作类型)
    '                       If Nvl(rsTmp!操作类型) = "微生物" Then mblnNotAddNew = True '微生物只能开一个检验项目
                vsExt.TextMatrix(vsExt.Rows - 1, 3) = 1  '子项
            End If
            
            rsTmp.MoveNext
        Loop
    End If
    
    '下一输入行
    If vsExt.RowData(vsExt.Rows - 1) <> 0 And Not mblnNotAddNew Then vsExt.AddItem ""
    vsExt.Row = vsExt.Rows - 1: vsExt.Col = 2
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsExt_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim lngRow As Long, lngCol As Long
    Dim strTip As String
    
    If mintType = 0 Then
        lngRow = vsExt.MouseRow: lngCol = vsExt.MouseCol
        If Between(lngRow, 0, vsExt.Rows - 1) And Between(lngCol, 0, vsExt.Cols - 1) Then
            If vsExt.Cell(flexcpPicture, lngRow, lngCol) Is Nothing Then
                If Me.TextWidth(vsExt.TextMatrix(lngRow, lngCol)) > vsExt.ColWidth(lngCol) - 15 Then
                    strTip = vsExt.TextMatrix(lngRow, lngCol)
                End If
            Else
                If Me.TextWidth(vsExt.TextMatrix(lngRow, lngCol)) > vsExt.ColWidth(lngCol) - 15 - 240 Then
                    strTip = vsExt.TextMatrix(lngRow, lngCol)
                End If
            End If
        End If
        vsExt.ToolTipText = strTip
    End If
End Sub

Private Sub vsExt_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If mintType = 0 Then
        If vsExt.Col = 1 And vsExt.MouseCol = 1 Then
            If x <= vsExt.CellLeft + 250 Then
                Call vsExt_KeyPress(vbKeySpace)
            End If
        End If
    End If
End Sub

Private Sub vsExt_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsExt.EditSelStart = 0
    vsExt.EditSelLength = zlCommFun.ActualLen(vsExt.EditText)
End Sub

Private Sub vsExt_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
'功能：限制某些列不允许编辑(该事件后于BeforeEdit,在EditText赋值之前)
    mblnReturn = False
        
    If mintType = 0 Then
        '只允许选择检查方法
        If Col <> 2 Then Cancel = True
    ElseIf mintType = 1 Or mintType = 4 Then
        '只允许编辑附加手术
        If cmd.Visible Then cmd.Visible = False '开始编辑了则隐藏按钮
        If Col <> 0 And mintType = 1 Or Col <> 2 And Col <> 0 And mintType = 4 Then Cancel = True
        '如果开启了新版LIS的组合项目模式则子项不允许输入
        If mblnNewLIS And mintType = 4 And Col = 2 Then
            If vsExt.TextMatrix(Row, 3) = "1" Then Cancel = True
        ElseIf mblnNewLIS And mintType = 4 And Col = 0 Then
            If Val(vsExt.TextMatrix(Row, 3)) = 0 Then Cancel = True
        End If
    End If
End Sub

Private Sub vsMethod_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If NewCol = 0 And NewRow <> -1 Then
        If vsMethod.TextMatrix(NewRow, 0) = "" Then
            Cancel = True
            vsMethod.Col = 1
        End If
    End If
End Sub

Private Sub vsMethod_Click()
    If fraMethod.Visible And vsMethod.Row >= 0 And vsMethod.Col >= 0 Then Call vsMethod_KeyPress(vbKeySpace)
End Sub

Private Sub vsMethod_KeyPress(KeyAscii As Integer)
    Dim strMethod As String
    Dim i As Long, j As Long
    Dim blnDo As Boolean
    
    With vsMethod
        If KeyAscii = 13 Then
            '检查方法的确认
            For i = 0 To .Rows - 1
                If .Cell(flexcpData, i, 0) = 1 Then
                    strMethod = strMethod & "," & .TextMatrix(i, 1)
                End If
            Next
            If strMethod = "" Then Exit Sub
            vsExt.TextMatrix(vsExt.Row, 2) = Mid(strMethod, 2)
            vsExt.Cell(flexcpData, vsExt.Row, 1) = 1 '方法设置后，自动选中该部位
            Set vsExt.Cell(flexcpPicture, vsExt.Row, 1) = img16.ListImages("c1").Picture
            
            fraMethod.Visible = False
            vsExt.SetFocus
        ElseIf KeyAscii = vbKeySpace Then
            '检查方法的选择与取消
            If .Cell(flexcpData, .Row, 0) = 1 Then
                '单选项目前也允许取消选择
                .Cell(flexcpData, .Row, 0) = 0
                Set .Cell(flexcpPicture, .Row, IIF(.RowData(.Row) = 3, 1, 0), .Row, 1) = img16.ListImages(IIF(.RowData(.Row) = 1, "o0", "c0")).Picture
                '同时取消该单选项的子项
                If .RowData(.Row) = 1 Then
                    For i = .Row + 1 To .Rows - 1
                        If .RowData(i) = 3 Then
                            If .Cell(flexcpData, i, 0) = 1 Then
                                .Cell(flexcpData, i, 0) = 0
                                Set .Cell(flexcpPicture, i, 1) = img16.ListImages("c0").Picture
                            End If
                        Else
                            Exit For
                        End If
                    Next
                End If
            Else
                blnDo = True
                If .RowData(.Row) = 3 Then
                    '主项没有选择时,子项不能选择
                    For i = .Row - 1 To 0 Step -1
                        If .RowData(i) <> 3 Then
                            If .Cell(flexcpData, i, 0) = 0 Then blnDo = False
                            Exit For
                        End If
                    Next
                End If
                If blnDo Then
                    .Cell(flexcpData, .Row, 0) = 1
                    Set .Cell(flexcpPicture, .Row, IIF(.RowData(.Row) = 3, 1, 0), .Row, 1) = img16.ListImages(IIF(.RowData(.Row) = 1, "o1", "c1")).Picture
                    If .RowData(.Row) = 1 Then '单选项选中时，取消其他单选项
                        For i = 0 To .Rows - 1
                            If i <> .Row And .RowData(i) = 1 Then
                                .Cell(flexcpData, i, 0) = 0
                                Set .Cell(flexcpPicture, i, 0, i, 1) = img16.ListImages("o0").Picture
                                For j = i + 1 To .Rows - 1 '同时取消该单选项的子项
                                    If .RowData(j) = 3 Then
                                        If .Cell(flexcpData, j, 0) = 1 Then
                                            .Cell(flexcpData, j, 0) = 0
                                            Set .Cell(flexcpPicture, j, 1) = img16.ListImages("c0").Picture
                                        End If
                                    Else
                                        Exit For
                                    End If
                                Next
                            End If
                        Next
                    End If
                End If
            End If
        End If
    End With
End Sub
