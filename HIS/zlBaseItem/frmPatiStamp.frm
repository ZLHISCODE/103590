VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Begin VB.Form frmPatiStamp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "病人标记设置"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10095
   Icon            =   "frmPatiStamp.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   10095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraUnit 
      Height          =   1815
      Left            =   1800
      TabIndex        =   0
      Top             =   840
      Width           =   3615
      Begin VB.TextBox txtName 
         Height          =   300
         Left            =   960
         MaxLength       =   10
         TabIndex        =   1
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtDays 
         Height          =   300
         Left            =   960
         MaxLength       =   3
         TabIndex        =   2
         Top             =   720
         Width           =   855
      End
      Begin VB.CheckBox chkSpecial 
         Caption         =   "应用于特殊病人图标设置"
         Height          =   225
         Left            =   120
         TabIndex        =   3
         Top             =   1170
         Width           =   3045
      End
      Begin VB.Label lblSet 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "标记名称"
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   300
         Width           =   720
      End
      Begin VB.Label lblSet 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0表示永久有效"
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   2
         Left            =   2280
         TabIndex        =   6
         Top             =   780
         Width           =   1170
      End
      Begin VB.Label lblSet 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "天"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   3
         Left            =   1920
         TabIndex        =   5
         Top             =   780
         Width           =   180
      End
      Begin VB.Label lblSet 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "有效天数"
         Height          =   180
         Index           =   4
         Left            =   120
         TabIndex        =   4
         Top             =   780
         Width           =   720
      End
   End
   Begin VB.ComboBox cboUnit 
      Height          =   300
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   0
      Width           =   1905
   End
   Begin VB.Frame fraInfo 
      Height          =   4575
      Left            =   5520
      TabIndex        =   11
      Top             =   945
      Width           =   3975
      Begin VB.PictureBox picBack 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   2295
         Left            =   1080
         ScaleHeight     =   2265
         ScaleWidth      =   2625
         TabIndex        =   12
         Top             =   1560
         Visible         =   0   'False
         Width           =   2655
         Begin VB.VScrollBar HScr 
            Height          =   2295
            LargeChange     =   50
            Left            =   2400
            Max             =   100
            SmallChange     =   100
            TabIndex        =   17
            Top             =   0
            Width           =   255
         End
         Begin VB.PictureBox pic标记 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1335
            Left            =   360
            ScaleHeight     =   1335
            ScaleWidth      =   1335
            TabIndex        =   13
            Top             =   120
            Width           =   1335
            Begin VB.PictureBox picIcon 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   615
               Index           =   0
               Left            =   120
               ScaleHeight     =   615
               ScaleWidth      =   615
               TabIndex        =   14
               Top             =   120
               Width           =   615
               Begin VB.Image imgICon 
                  Height          =   360
                  Index           =   0
                  Left            =   120
                  Picture         =   "frmPatiStamp.frx":6852
                  Top             =   0
                  Width           =   360
               End
               Begin VB.Label lblSelect 
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  ForeColor       =   &H80000008&
                  Height          =   360
                  Index           =   0
                  Left            =   120
                  TabIndex        =   16
                  Top             =   120
                  Width           =   300
               End
               Begin VB.Label lblInfo 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  Caption         =   "PDA"
                  ForeColor       =   &H80000008&
                  Height          =   180
                  Index           =   0
                  Left            =   120
                  TabIndex        =   15
                  Top             =   450
                  Width           =   270
               End
            End
         End
      End
      Begin VB.TextBox txtInfo 
         Height          =   300
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   21
         Top             =   1200
         Width           =   1935
      End
      Begin VB.ComboBox cbo标记 
         Height          =   300
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   240
         Width           =   1905
      End
      Begin VB.CommandButton cmdImage 
         Appearance      =   0  'Flat
         Caption         =   "&P"
         Height          =   300
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         ToolTipText     =   "选择项目(F4)"
         Top             =   720
         Width           =   270
      End
      Begin MSComctlLib.ImageCombo imaCustom 
         Height          =   315
         Left            =   1080
         TabIndex        =   20
         Top             =   720
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   556
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Locked          =   -1  'True
      End
      Begin VB.Label lblSet 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "标记图形"
         Height          =   180
         Index           =   7
         Left            =   240
         TabIndex        =   24
         Top             =   780
         Width           =   720
      End
      Begin VB.Label lblSet 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "个性标记"
         Height          =   180
         Index           =   9
         Left            =   240
         TabIndex        =   23
         Top             =   300
         Width           =   720
      End
      Begin VB.Label lblSet 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "标记说明"
         Height          =   180
         Index           =   8
         Left            =   240
         TabIndex        =   22
         Top             =   1260
         Width           =   720
      End
   End
   Begin VB.Frame fraLine 
      BorderStyle     =   0  'None
      Height          =   3615
      Left            =   5280
      TabIndex        =   10
      Top             =   960
      Width           =   100
   End
   Begin VB.Frame fraUd 
      Height          =   3855
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Width           =   4935
      Begin XtremeReportControl.ReportControl UnitReportControl 
         Height          =   2415
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   3495
         _Version        =   589884
         _ExtentX        =   6165
         _ExtentY        =   4260
         _StockProps     =   0
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfPrint 
      Height          =   420
      Left            =   240
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   5400
      Visible         =   0   'False
      Width           =   855
      _cx             =   1508
      _cy             =   741
      Appearance      =   2
      BorderStyle     =   1
      Enabled         =   0   'False
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
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
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
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   2000
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
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   27
      Top             =   6135
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPatiStamp.frx":6F54
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15372
         EndProperty
      EndProperty
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
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   1680
      Top             =   240
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmPatiStamp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const FSHIFT = 4
Const FCONTROL = 8
Const FALT = 16

Const VK_DELETE = &H2E
Const VK_F1 = &H70
Const VK_F5 = &H74
Const VK_INSERT = &H2D

Const conMenu_FilePopup = 1    '文件
Const conMenu_EditPopup = 3    '编辑
Const conMenu_ViewPopup = 7    '查看
Const conMenu_ToolPopup = 8    '工具
Const conMenu_HelpPopup = 9    '帮助
Const conMenu_File_PrintSet = 101        '*打印设置(&S)…
Const conMenu_File_Preview = 102         '*预览(&V)
Const conMenu_File_Print = 103           '*打印(&P)
Const conMenu_File_Excel = 104           '输出到&Excel…
Const conMenu_Edit_Save = 3091        '*保存
Const conMenu_File_Exit = 191            '*退出(&X)
Const conMenu_Edit_Reuse = 3009      '*启用(&U)
Const conMenu_Edit_FileMan = 3047
Const conMenu_Edit_NewParent = 3051   '*新分类(&N)
Const conMenu_Edit_ModifyParent = 3053    '*修改分类(&M)
Const conMenu_Edit_DeleteParent = 3054    '*删除分类(&D)
Const conMenu_Edit_Leave_Add = 3561    '增加
Const conMenu_Edit_NewItem = 3001    '*新项目(&A)
Const conMenu_Edit_Modify = 3003     '*修改(&M)
Const conMenu_Edit_Delete = 3004     '*删除(&D)
Const conMenu_View_ToolBar = 701              '工具栏(&T)
Const conMenu_View_ToolBar_Button = 7011         '标准按钮(&S)
Const conMenu_View_ToolBar_Text = 7012           '文本标签(&T)
Const conMenu_View_ToolBar_Size = 7013           '大图标(&B)
Const conMenu_View_StatusBar = 702            '状态栏(&S)
Const conMenu_View_Refresh = 791              '*刷新(&R)
Const conMenu_Help_Help = 901        '*帮助主题(&H)
Const conMenu_Help_Web = 902         '&WEB上的中联
Const conMenu_Help_Web_Home = 9021       '中联主页(&H)
Const conMenu_Help_Web_Forum = 9023      '中联论坛(&F)
Const conMenu_Help_Web_Mail = 9022       '*发送反馈(&M)
Const conMenu_Help_About = 991       '关于(&A)…
Const conMenu_View_Find = 721

Const madLongVarCharDefault As Integer = 10          '字符型字段缺省长度
Const madDoubleDefault As Integer = 18               '数字型字段缺省长度
Const madDbDateDefault As Integer = 20               '日期型字段缺省长度

Const COL_NULL = 0
Const COL_标注 = 1
Const COL_说明 = 2
Const COL_主题序号 = 3
Const COL_有效天数 = 4
Const COL_原始主题 = 5
Const COL_原始标记 = 6
Const COL_主题说明 = 7
Const COL_是否特殊 = 8
  
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private mRect As RECT

Private Type TYPE_UNIT
    病区ID  As Long
    主题序号 As Long
    标记序号 As Long
    说明 As String
    图形索引 As Long
    有效天数 As Long
    原始主题 As Long
    原始标记 As Long
End Type

Private mUnit As TYPE_UNIT

Const Enable_Color = &HE0E0E0
Const UnEnable_Color = &H80000005

Private mblnChange As Boolean '记录标记内容变动
Private mstrSubject As String '标记分类名称
Private mlngDay As Long '标记分类天数
Private mintSpecial As Integer '标记分类是否应用与特殊人群
Private mLngCount As Long  '存放标记分类数目

Private m病区ID As Long
Private mstr病区名称 As String

Private mcbrToolBars As CommandBar  '工具栏
Private mcbrMenuBars As CommandBarControl
Const mlngImgIndex As Long = 0 '定义图片索引从第几个开始显示

Private mblnOK As Boolean
Private mrsData As New ADODB.Recordset

Public Function ShowMe(ByVal FrmParent As Form) As Boolean
    mblnOK = False
    Me.Show 1, FrmParent
    ShowMe = mblnOK
End Function

Private Sub cboUnit_Click()
    If cboUnit.ListCount > 0 And m病区ID <> Val(cboUnit.ItemData(cboUnit.ListIndex)) Then
        m病区ID = Val(cboUnit.ItemData(cboUnit.ListIndex))
        mstr病区名称 = cboUnit.Text
    
        Call RefreshData
    End If
End Sub

Private Sub cboUnit_KeyPress(KeyAscii As Integer)
    Call zlControl.CboMatchIndex(cboUnit.hwnd, KeyAscii)
End Sub

Private Sub cbo标记_Click()
'-------------------------------------------------
'功能:根据选择主题序号改变标记内容位置
'-------------------------------------------------
    Dim strTag As String
    Dim lngPreID As Long
    Dim rsTemp As New ADODB.Recordset
    Dim strSql As String
    Dim lngRowIndex As Long, lngRow As Long, lngOldID As Long
    Dim strFileds As String, strValues As String
    Dim str标记 As String, strCaption As String
    Dim intDay As Integer, intSpecial As Integer
    
    If UnitReportControl.Records.Count = 0 Then Exit Sub
    If cbo标记.ListIndex = -1 Or fraInfo.Tag = "新增" Or mblnChange = False Then Exit Sub
    If UnitReportControl.FocusedRow.GroupRow And UnitReportControl.FocusedRow.Childs.Count <> 0 Then Exit Sub
    If mrsData Is Nothing Then Exit Sub
    
    strFileds = "主题序号," & adDouble & ",18|标记序号," & adDouble & ",18|说明," & adLongVarChar & ",100|图形索引," & _
        adDouble & ",18|有效天数," & adDouble & ",18|是否特殊," & adInteger & ",1|原始主题序号," & adDouble & ",18|原始标记序号," & adDouble & ",18"
    Call Record_Init(rsTemp, strFileds)
    'A.主题序号,A.标记序号,A.说明,A.图形索引,A.有效天数,A.是否特殊,A.主题序号 原始主题序号,A.标记序号 原始标记序号
    strFileds = "主题序号|标记序号|说明|图形索引|有效天数|是否特殊|原始主题序号|原始标记序号"
    
    lngRowIndex = UnitReportControl.FocusedRow.Index
    
    str标记 = ""
    mrsData.Filter = ""
    For lngRow = 0 To UnitReportControl.Rows.Count - 1
        If Not UnitReportControl.Rows(lngRow).GroupRow Then
            lngOldID = Val(Split(UnitReportControl.Rows(lngRow).Record(COL_主题序号).Record.Tag, "-")(0))
            mrsData.Filter = "主题序号=" & lngOldID & " and 标记序号=0"
            If mrsData.RecordCount > 0 Then
                strCaption = Nvl(mrsData!说明)
                intDay = Val(Nvl(mrsData!有效天数))
                intSpecial = Val(Nvl(mrsData!是否特殊))
            End If
            
            If UnitReportControl.Rows(lngRow).Index = lngRowIndex Then
                mUnit.主题序号 = Val(cbo标记.ItemData(cbo标记.ListIndex))
                lngPreID = AgainComputePreId(Val(cbo标记.ItemData(cbo标记.ListIndex))) '获取标记序号
                mUnit.标记序号 = lngPreID
                
                mrsData.Filter = "主题序号=" & mUnit.主题序号 & " and 标记序号=0"
                If mrsData.RecordCount > 0 Then
                    mUnit.有效天数 = Val(Nvl(mrsData!有效天数))
                End If
                str标记 = mUnit.主题序号 & "-" & mUnit.标记序号 & "-" & m病区ID & "-" & mUnit.有效天数
            Else
                mUnit.主题序号 = Val(Split(UnitReportControl.Rows(lngRow).Record(COL_主题序号).Record.Tag, "-")(0))
                mUnit.标记序号 = Val(Split(UnitReportControl.Rows(lngRow).Record(COL_主题序号).Record.Tag, "-")(1))
                mUnit.有效天数 = intDay 'Val(zlCommFun.Nvl(UnitReportControl.Rows(lngRow).Record(COL_有效天数).Value, 0))
            End If
                        
            mUnit.说明 = zlCommFun.Nvl(UnitReportControl.Rows(lngRow).Record(COL_说明).Value)
            mUnit.图形索引 = Val(zlCommFun.Nvl(UnitReportControl.Rows(lngRow).Record(COL_标注).Icon, 0))
            mUnit.原始主题 = zlCommFun.Nvl(UnitReportControl.Rows(lngRow).Record(COL_原始主题).Value, 0)
            mUnit.原始标记 = zlCommFun.Nvl(UnitReportControl.Rows(lngRow).Record(COL_原始标记).Value, 0)
            If mUnit.主题序号 <> mUnit.原始主题 Then '主题序号变化时，则检查是否已经使用
                If CheckUseUnit(m病区ID, mUnit.原始主题, mUnit.原始标记) Then
                    Call zlControl.CboLocate(cbo标记, lngOldID, True)
                    Exit Sub
                End If
            End If
            '检查主题序号是否存在 不存在就添加
            rsTemp.Filter = "主题序号=" & lngOldID & " and 标记序号=0"
            If rsTemp.RecordCount = 0 Then
                strValues = lngOldID & "|" & 0 & "|" & strCaption & "|0|" & _
                    intDay & "|" & intSpecial & "|" & mUnit.原始主题 & "|" & mUnit.原始标记
                Call Record_Add(rsTemp, strFileds, strValues)
            End If
            If Val(Split(UnitReportControl.Rows(lngRow).Record(COL_主题序号).Record.Tag, "-")(1)) <> 0 Then
                strValues = mUnit.主题序号 & "|" & mUnit.标记序号 & "|" & mUnit.说明 & "|" & mUnit.图形索引 & "|" & _
                    mUnit.有效天数 & "|0|" & mUnit.原始主题 & "|" & mUnit.原始标记
                Call Record_Add(rsTemp, strFileds, strValues)
            End If
        End If
    Next lngRow
    
    rsTemp.Filter = 0
    rsTemp.Sort = "主题序号,标记序号"
    'Call OutputRsData(rsTemp)
    Call RefreshData(0, str标记, rsTemp)
    mblnChange = True
'    With UnitReportControl.FocusedRow.Record(COL_主题序号)
'        .GroupCaption = "分组：" & cbo标记.ItemData(cbo标记.ListIndex) & "-" & cbo标记.Text
'        strTag = .Record.Tag
'        lngPreID = AgainComputePreId(Val(cbo标记.ItemData(cbo标记.ListIndex))) '获取标记序号
'        .Record.Tag = cbo标记.ItemData(cbo标记.ListIndex) & "-" & lngPreID & "-" & Split(strTag, "-")(2)
'    End With
'
'    UnitReportControl.Populate

End Sub

Private Sub cbo标记_GotFocus()
    If picBack.Visible = True Then
        picBack.Visible = False
        cmdImage.Enabled = True
    End If
End Sub

Private Sub cbo标记_KeyPress(KeyAscii As Integer)
    Call zlControl.CboMatchIndex(cbo标记.hwnd, KeyAscii)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub cbsMain_Resize()
    Call ResizeState
End Sub

Private Sub cmdImage_Click()
'功能显示现有图片信息
    Dim i As Integer, j As Integer
    Dim lngCurXCount As Long
    Dim lngH As Integer, lngW As Integer '记录picture的高度和宽度
    Dim lngX1 As Long 'pictrue之间的间隔
    Dim lngX As Long, lngY As Long  '设定image的顶部和左侧边距
    Dim lngIndex As Long
    Dim vRect As RECT
    Dim vRect1 As RECT
    
    
    lngIndex = 0
    lngY = 60
    lngX = 60

    imgICon(lngIndex).Top = lngY
    imgICon(lngIndex).Left = lngX
    
    lblSelect(lngIndex).Top = lngY / 2
    lblSelect(lngIndex).Left = lngX / 2
    lblSelect(lngIndex).Width = imgICon(lngIndex).Width + lngX
    lblSelect(lngIndex).Height = imgICon(lngIndex).Height + lngY
    
    lblInfo(lngIndex).FontSize = 8
    lblInfo(lngIndex).Top = lngY + imgICon(lngIndex).Width + lngY / 2
    lblInfo(lngIndex).Caption = zlCommFun.GetPaitSignImageList(1).ListImages(mlngImgIndex + 1).Key
    
    picIcon(lngIndex).Top = 0
    picIcon(lngIndex).Left = 0
    picIcon(lngIndex).Height = imgICon(lngIndex).Height + lngY + lngY / 2 + lblInfo(lngIndex).Height + 10
    picIcon(lngIndex).Width = imgICon(lngIndex).Width + imgICon(lngIndex).Left * 2 + lngX / 2
    
    lngH = picIcon(lngIndex).Height
    lngW = picIcon(lngIndex).Width
    
    lblInfo(lngIndex).Left = (lngW - lblInfo(lngIndex).Width) / 2
    
    '获取计算picback的位置的宽度
    vRect = zlControl.GetControlRect(imaCustom.hwnd)
    vRect1 = zlControl.GetControlRect(fraInfo.hwnd)
    picBack.Top = vRect.Bottom - vRect1.Top
    picBack.Left = vRect.Left - vRect1.Left
    picBack.Width = vRect1.Right - vRect.Left - 10
    
    pic标记.Width = picBack.ScaleWidth - HScr.Width
    
    '计算每行可存放的图片数量
    lngCurXCount = (pic标记.Width - HScr.Width) \ lngW
    '重新计算位置
    lngX1 = (pic标记.Width - HScr.Width - (lngW * lngCurXCount)) / (lngCurXCount + 1)
    picIcon(lngIndex).Left = lngX1
    
    imgICon(lngIndex).Picture = zlCommFun.GetPaitSignImageList(1).ListImages(mlngImgIndex + 1).Picture
    
    HScr.Top = 0
    HScr.Min = 0
    HScr.Left = pic标记.Width
    HScr.Value = 0
    HScr.Height = picBack.ScaleHeight
    
    picBack.Visible = True
    picBack.ZOrder 0
    pic标记.Visible = True
    pic标记.Top = 0
    pic标记.Left = 0
    pic标记.SetFocus
    
    For i = 1 To picIcon.Count - 1
        If i < lngCurXCount Then
            picIcon(i).Top = 0
            picIcon(i).Left = lngW * i + (i + 1) * lngX1
        Else
            picIcon(i).Top = lngH * ((i \ lngCurXCount))
            picIcon(i).Left = lngW * (i Mod lngCurXCount) + ((i Mod lngCurXCount) + 1) * lngX1
        End If
        picIcon(i).Width = picIcon(lngIndex).Width
        picIcon(i).Height = picIcon(lngIndex).Height
        
        imgICon(i).Top = imgICon(lngIndex).Top
        imgICon(i).Left = imgICon(lngIndex).Left
        
        lblSelect(i).Top = lblSelect(lngIndex).Top
        lblSelect(i).Left = lblSelect(lngIndex).Left
        lblSelect(i).Width = lblSelect(lngIndex).Width
        lblSelect(i).Height = lblSelect(lngIndex).Height
        
        lblInfo(i).FontSize = lblInfo(lngIndex).FontSize
        lblInfo(i).Top = lblInfo(lngIndex).Top
        lblInfo(i).Left = (lngW - lblInfo(i).Width) / 2
    Next i
    
    pic标记.Height = picIcon(i - 1).Top + picIcon(i - 1).Height
    pic标记.Refresh
    
    If pic标记.ScaleHeight - picBack.ScaleHeight <= 0 Then
        HScr.Max = 0
        HScr.Min = 0
        HScr.Visible = False
    Else
        HScr.Max = pic标记.ScaleHeight - picBack.ScaleHeight
        HScr.Visible = True
    End If
    cmdImage.Enabled = False
    
    If Not imaCustom.SelectedItem Is Nothing Then
        lngIndex = imaCustom.SelectedItem.Index
        If lngIndex > 0 And lngIndex <= picIcon.Count Then
            If HScr.Max > 0 Then
                '标记区域小于图片的位置，说明图片显示不完
                If picBack.ScaleHeight < picIcon(lngIndex - 1).Top + picIcon(lngIndex - 1).Height Then
                    If picIcon(lngIndex - 1).Top + picIcon(lngIndex - 1).Height - picBack.ScaleHeight > HScr.Max Then
                        HScr.Value = HScr.Max
                    Else
                        HScr.Value = picIcon(lngIndex - 1).Top + picIcon(lngIndex - 1).Height - picBack.ScaleHeight
                    End If
                End If
            End If
            Call ShowSelect(lngIndex - 1)
        End If
    End If
End Sub

Private Sub LoadICon()
'加载自定义图标
    Dim i As Integer, j As Integer
    On Error GoTo ErrHand
    i = 1
    For j = mlngImgIndex + 1 To zlCommFun.GetPaitSignImageList(1).ListImages.Count - 1
        Load picIcon(i)
        picIcon(i).Visible = True
        
        '加载图片信息
        Load imgICon(i)
        imgICon(i).Visible = True
        Set imgICon(i).Container = picIcon(i)
        imgICon(i).Picture = zlCommFun.GetPaitSignImageList(1).ListImages(j + 1).Picture
        
        '加载选择控件
        Load lblSelect(i)
        lblSelect(i).Visible = True
        Set lblSelect(i).Container = picIcon(i)
        
        '加载图片说明
        Load lblInfo(i)
        lblInfo(i).Visible = True
        Set lblInfo(i).Container = picIcon(i)
        lblInfo(i).Caption = zlCommFun.GetPaitSignImageList(1).ListImages(j + 1).Key
        
        i = i + 1
    Next j
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function GetMarkCount() As Long
    '获取标记项目总数
    Dim lngRow As Long
    Dim lngCount As Long
    
    For lngRow = 0 To UnitReportControl.Rows.Count - 1
        '标记序号=0的为标记主题分类，不进行统计
        If Not UnitReportControl.Rows(lngRow).GroupRow And UnitReportControl.Rows(lngRow).Childs.Count = 0 Then
            If Val(Split(UnitReportControl.Rows(lngRow).Record(COL_主题序号).Record.Tag, "-")(1)) <> 0 Then
                lngCount = lngCount + 1
            End If
        End If
    Next lngRow
    
    GetMarkCount = lngCount
End Function

Private Sub RefreshStateInfo()
'------------------------------------------------------------------------------------------------------------------
'功能：刷新状态栏显示信息
'-----------------------------------------------------------------------------------------------------------------
    stbThis.Panels(2).Text = "共有 " & GetMarkCount & " 个标记内容！"
End Sub

Private Sub UnLoadImage()
'功能:卸载pic标注上的所有控件
    Dim i As Integer
    For i = picIcon.Count - 1 To 1 Step -1
        Unload imgICon(i)
        Unload lblInfo(i)
        Unload lblSelect(i)
        Unload picIcon(i)
    Next i
    picBack.Visible = False
    cmdImage.Enabled = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 39 Then KeyCode = 0
    If KeyCode = 27 And picBack.Visible = True Then
        picBack.Visible = False
        cmdImage.Enabled = True
    End If
End Sub

Private Sub Form_Load()
    Call RestoreWinState(Me, App.ProductName)
    '加载菜单工具栏
    Call InitCommandBar
    '提取病区信息
    Call InitUnits
    '加载主题标致信息
    Call InitReportControl
    '读取数据
    Call RefreshData
End Sub

Private Sub AddImage()
'------------------------------------
'功能:加载所有图片信息到ImageCombo
'------------------------------------
    Dim objNewItem As ComboItem
    Dim i As Long
 
    imaCustom.ImageList = zlCommFun.GetPaitSignImageList(0)
    For i = 1 To zlCommFun.GetPaitSignImageList(0).ListImages.Count - mlngImgIndex
        Set objNewItem = imaCustom.ComboItems.Add(i, "A" & i, zlCommFun.GetPaitSignImageList(0).ListImages(mlngImgIndex + i).Key, mlngImgIndex + i)
    Next i
    
End Sub

Public Sub zlRptPrint(ByVal bytMode As Byte)
    '功能:将数据复制到可打印的对象，调用打印
    '参数:  bytMode，1-打印;2-预览;3-输出到EXCEL
    If UnitReportControl.Records.Count = 0 Then Exit Sub
    
    '-------------------------------------------------
    '复制数据表格
    If zlReportToVSFlexGrid(vsfPrint, UnitReportControl) = False Then Exit Sub
    
    '-------------------------------------------------
    '调用打印部件处理
    Dim objPrint As New zlPrint1Grd
    Dim objAppRow As zlTabAppRow
    
    Set objPrint.Body = vsfPrint
    
    objPrint.Title.Text = "病区标记内容清单"
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

Private Sub InitCommandBar()
'功能:初始化菜单栏
    Dim cbrTools As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim objCustom As CommandBarControlCustom
    Dim objControl As CommandBarControl
    Dim strProductName As String
    On Error GoTo ErrHand
    
    strProductName = GetSetting("ZLSOFT", "注册信息", "产品名称", "中联")
    
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    With cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .ShowTextBelowIcons = False
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
        .UseSharedImageList = False '显示图形
    End With
    
        '菜单定义
    cbsMain.ActiveMenuBar.Title = "菜单栏"
    cbsMain.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    Set mcbrMenuBars = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "文件(&F)", -1, False)
    mcbrMenuBars.ID = conMenu_FilePopup
    With mcbrMenuBars.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_File_PrintSet, "打印设置(&S)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Preview, "预览(&V)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Print, "打印(&P)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Excel, "输出到&Excel")
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "保存(&S)")
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "取消(&Z)")
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出(&X)")
        cbrControl.BeginGroup = True
    End With

    Set mcbrMenuBars = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_EditPopup, "编辑(&E)", -1, False)
    mcbrMenuBars.ID = conMenu_EditPopup
    With mcbrMenuBars.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewParent, "新增分类(&I)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_ModifyParent, "修改分类(&U) ")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_DeleteParent, "删除分类(&E)")
    
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "新增(&A)")
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Modify, "修改(&M)")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Delete, "删除(&D)")
    End With

    Set mcbrMenuBars = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "查看(&V)", -1, False)
    mcbrMenuBars.ID = conMenu_ViewPopup
    With mcbrMenuBars.CommandBar.Controls
        Set cbrControl = .Add(xtpControlPopup, conMenu_View_ToolBar, "工具栏(&T)")
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Button, "标准按钮(&S)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_View_ToolBar_Text, "文本标签(&T)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_View_StatusBar, "状态栏(&S)")
        Set cbrControl = .Add(xtpControlButton, conMenu_View_Refresh, "刷新(&R)"): cbrControl.BeginGroup = True
    End With
    
    Set mcbrMenuBars = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "帮助(&H)", -1, False)
    mcbrMenuBars.ID = conMenu_HelpPopup
    With mcbrMenuBars.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助主题(&H)")
        Set cbrControl = .Add(xtpControlPopup, conMenu_Help_Web, "&WEB上的" & strProductName)
        
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Home, strProductName & "主页(&H)", -1, False
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Forum, strProductName & "论坛(&F)", -1, False  '固有
        cbrControl.CommandBar.Controls.Add xtpControlButton, conMenu_Help_Web_Mail, "发送反馈(&M)", -1, False
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_About, "关于(&A)..."): cbrControl.BeginGroup = True
    End With
    
     '快键绑定
    With cbsMain.KeyBindings
        .Add FCONTROL, Asc("P"), conMenu_File_Print
        .Add FCONTROL, Asc("A"), conMenu_Edit_NewItem
        .Add FCONTROL, Asc("M"), conMenu_Edit_Modify
        .Add FCONTROL, Asc("S"), conMenu_Edit_Save
        .Add FCONTROL, Asc("Z"), conMenu_Edit_Reuse
        .Add FSHIFT, VK_INSERT, conMenu_Edit_NewParent
        .Add FSHIFT, VK_DELETE, conMenu_Edit_DeleteParent
        .Add 0, VK_DELETE, conMenu_Edit_Delete
        .Add 0, VK_F5, conMenu_View_Refresh
        .Add 0, VK_F1, conMenu_Help_Help
    End With
    
    '--添加工具栏
    Set mcbrToolBars = cbsMain.Add("工具栏", xtpBarTop)
    mcbrToolBars.EnableDocking xtpFlagStretched
    With mcbrToolBars.Controls
        Set cbrTools = .Add(xtpControlPopup, conMenu_Edit_FileMan, "分类", -1, False)
        cbrTools.IconId = conMenu_Edit_FileMan
        cbrTools.ToolTipText = "标记分类"
        cbrTools.BeginGroup = True
        
        cbrTools.CommandBar.Controls.Add xtpControlButton, conMenu_Edit_NewParent, "新增"
        cbrTools.CommandBar.Controls.Add xtpControlButton, conMenu_Edit_ModifyParent, "修改"
        cbrTools.CommandBar.Controls.Add xtpControlButton, conMenu_Edit_DeleteParent, "删除"
        
        Set cbrTools = .Add(xtpControlPopup, conMenu_Edit_Leave_Add, "标记", -1, False)
        cbrTools.IconId = conMenu_Edit_NewItem
        cbrTools.ToolTipText = "标记内容"
        
        cbrTools.CommandBar.Controls.Add xtpControlButton, conMenu_Edit_NewItem, "新增"
        cbrTools.CommandBar.Controls.Add xtpControlButton, conMenu_Edit_Modify, "修改"
        cbrTools.CommandBar.Controls.Add xtpControlButton, conMenu_Edit_Delete, "删除"
        

        Set cbrTools = .Add(xtpControlButton, conMenu_Edit_Save, "保存")
        cbrTools.ToolTipText = "保存"
        cbrTools.BeginGroup = True
        
        Set cbrTools = .Add(xtpControlButton, conMenu_Edit_Reuse, "取消")
        cbrTools.ToolTipText = "取消"

        Set cbrTools = .Add(xtpControlButton, conMenu_Help_Help, "帮助")
        cbrTools.ToolTipText = "帮助"
        cbrTools.BeginGroup = True
        Set cbrTools = .Add(xtpControlButton, conMenu_File_Exit, "退出")

    End With
    
    For Each cbrControl In mcbrToolBars.Controls
        cbrControl.Style = xtpButtonIconAndCaption
    Next
    
    '工具栏右侧病区下拉框选择
    With mcbrToolBars.Controls
        Set objControl = .Add(xtpControlLabel, conMenu_View_Find, "病区")
        objControl.flags = xtpFlagRightAlign
        Set objCustom = .Add(xtpControlCustom, conMenu_View_Find, "病区")
        objCustom.Handle = Me.cboUnit.hwnd
        objCustom.flags = xtpFlagRightAlign
        objControl.IconId = conMenu_View_Find
    End With
    
    '加载图片信息
    Call AddImage
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub InitReportControl()
'功能:初始化ReportControl

    Dim Column As ReportColumn
    
    With UnitReportControl
        Set Column = .Columns.Add(COL_NULL, " ", 10, False)
        Column.Editable = False: Column.Groupable = False: Column.Sortable = False: Column.Alignment = xtpAlignmentCenter
        Set Column = .Columns.Add(COL_标注, "标注", 50, True)
        Column.Editable = False: Column.Groupable = False: Column.AllowDrag = False
        
        Set Column = .Columns.Add(COL_说明, "说明", 190, True)
        Column.AllowDrag = False: Column.Editable = False: Column.Groupable = False
        Set Column = .Columns.Add(COL_主题序号, "主题序号", 0, False)
        Column.Visible = False: Column.Editable = False: Column.Groupable = True
        Set Column = .Columns.Add(COL_有效天数, "有效天数", 60, True)
        Column.AllowDrag = False: Column.Editable = False: Column.Groupable = False
        Set Column = .Columns.Add(COL_原始主题, "原始主题", 0, False)
        Column.Visible = False: Column.Editable = False: Column.Groupable = False
        Set Column = .Columns.Add(COL_原始标记, "原始标记", 0, False)
        Column.Visible = False: Column.Editable = False: Column.Groupable = False
        Set Column = .Columns.Add(COL_主题说明, "主题说明", 0, False)
        Column.Visible = False: Column.Editable = False: Column.Groupable = False
        Set Column = .Columns.Add(COL_是否特殊, "是否特殊", 0, False)
        Column.Visible = False: Column.Editable = False: Column.Groupable = False
        
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .ShadeGroupHeadings = False
            .NoItemsText = "没有可显示的标记分类和标记内容信息..."
        End With
        
        .AllowColumnResize = False
        .ShowItemsInGroups = False '是否按排序自分理处分组
        .PreviewMode = True
        .MultipleSelection = False '会引发SelectionChanged事件
        .SetImageList zlCommFun.GetPaitSignImageList(0)
            
        .GroupsOrder.Add .Columns(COL_主题序号)
        .GroupsOrder(0).SortAscending = True
        .GroupsOrder(0).Groupable = True
        
        '分组之后可能失去记录集中的顺序,因此强行加入排序列
        .SortOrder.Add .Columns(COL_说明)
        .SortOrder(0).SortAscending = True
        .SortOrder.Add .Columns(COL_主题序号)
        .SortOrder(1).SortAscending = True
    End With
    
    Call LoadICon
End Sub

Private Function RefreshData(Optional lngPreIdx As Long, Optional str标记 As String = "", Optional ByVal rsTemp As ADODB.Recordset) As Boolean
'-------------------------------------------------------------
'功能:提取病区个性化设置
'参数:lngPreIdx 选择行索引,str标记 选择行信息（用来快速定位）
'说明 lngPreIdx=-1时不进行病区标记分类检查
'-------------------------------------------------------------
    Dim strUnit As String, StrInfo As String, strDay As String, strOldUnit As String
    Dim lngImgIndex As Long
    Dim blnDouble As Boolean
    Dim lngIndex As Long '存放当前序号
    Dim blnRead As Boolean
    Dim strSql As String
    'Dim rsTemp As New ADODB.Recordset
    Dim strSubject As String '存放标记分类的信息
    Dim objRow As ReportRow, i As Long
    Dim strFileds As String, strValues As String
    
    mblnChange = False
    Screen.MousePointer = 11
    On Error GoTo ErrHand
    
    mLngCount = CheckUnitSubject(m病区ID)
    
    If rsTemp Is Nothing Then blnRead = True
    If blnRead = False Then
        If rsTemp.State = adStateClosed Then blnRead = True
    End If
    If blnRead = True Then
        
        strFileds = "主题序号," & adDouble & ",18|标记序号," & adDouble & ",18|说明," & adLongVarChar & ",100|图形索引," & _
            adDouble & ",18|有效天数," & adDouble & ",18,|是否特殊," & adInteger & ",1,|原始主题序号," & adDouble & ",18|原始标记序号," & adDouble & ",18"
        Call Record_Init(mrsData, strFileds)
        strFileds = "主题序号|标记序号|说明|图形索引|有效天数|是否特殊|原始主题序号|原始标记序号"
         '提取病区信息
        strSql = _
            " SELECT A.主题序号,A.标记序号,A.说明,A.图形索引,A.有效天数,A.是否特殊,A.主题序号 原始主题序号,A.标记序号 原始标记序号" & vbNewLine & _
            " FROM 病区标记内容 A,病区标记内容 B" & vbNewLine & _
            " WHERE  " & IIF(m病区ID = 0, " B.病区ID IS NULL ", " A.病区ID=B.病区ID ") & " And A.主题序号=B.主题序号 And B.标记序号=0 " & IIF(m病区ID = 0, " And A.病区ID IS NULL ", " And A.病区ID=[1] ") & vbNewLine & _
            " ORDER BY A.主题序号,A.标记序号"
                
        Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "提取病区主题信息", m病区ID)
    End If
    
    UnitReportControl.Records.DeleteAll
    
    If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
    With rsTemp
        Do While Not .EOF
            If zlCommFun.Nvl(!标记序号) = 0 Then
                If strSubject <> "" Then
                    strUnit = strSubject
                    StrInfo = "此分类下没有可显示的标记内容信息..."
                    lngImgIndex = 0
                    AddRecord strUnit, lngImgIndex, StrInfo, mlngDay, strOldUnit
                    strSubject = ""
                End If
                mstrSubject = zlCommFun.Nvl(!说明, "个性标注" & zlCommFun.Nvl(!主题序号))
                mlngDay = Val(zlCommFun.Nvl(!有效天数, 0))
                mintSpecial = Val(zlCommFun.Nvl(!是否特殊, 0))
                strSubject = zlCommFun.Nvl(!主题序号) & "-" & zlCommFun.Nvl(!标记序号) & "-" & m病区ID
                strOldUnit = zlCommFun.Nvl(!原始主题序号) & "-" & zlCommFun.Nvl(!原始标记序号) & "-" & m病区ID
            Else
                strUnit = zlCommFun.Nvl(!主题序号) & "-" & zlCommFun.Nvl(!标记序号) & "-" & m病区ID
                strOldUnit = zlCommFun.Nvl(!原始主题序号) & "-" & zlCommFun.Nvl(!原始标记序号) & "-" & m病区ID
                StrInfo = zlCommFun.Nvl(!说明)
                strDay = zlCommFun.Nvl(!有效天数, 0)
                lngImgIndex = zlCommFun.Nvl(!图形索引, 0)
                AddRecord strUnit, lngImgIndex, StrInfo, mlngDay, strOldUnit
                strSubject = ""
            End If
            If blnRead = True Then
                strValues = Val(zlCommFun.Nvl(!主题序号)) & "|" & Val(zlCommFun.Nvl(!标记序号)) & "|" & zlCommFun.Nvl(!说明) & "|" & Val(zlCommFun.Nvl(!图形索引)) & "|" & _
                   Val(zlCommFun.Nvl(!有效天数)) & "|" & Val(zlCommFun.Nvl(!是否特殊)) & "|" & Val(zlCommFun.Nvl(!原始主题序号)) & "|" & Val(zlCommFun.Nvl(!原始标记序号))
                Call Record_Add(mrsData, strFileds, strValues)
            End If
        .MoveNext
        Loop
    End With
    
    If strSubject <> "" Then
        strUnit = strSubject
        StrInfo = "此分类下没有可显示的标记内容信息..."
        lngImgIndex = 0
        AddRecord strUnit, lngImgIndex, StrInfo, mlngDay, strOldUnit
        strSubject = ""
    End If
    
    UnitReportControl.Populate
    
    If UnitReportControl.Rows.Count <> 0 Then
        Call UnitRefresh(lngPreIdx, str标记)
    Else
        Call SetFraResize(True)
        txtName.Enabled = False
        txtName.Text = ""
        txtDays.Enabled = False
        txtDays.Text = ""
        txtName.BackColor = Enable_Color
        txtDays.BackColor = Enable_Color
        chkSpecial.Enabled = False
        chkSpecial.Value = 0
        chkSpecial.Visible = (m病区ID = 0)
    End If
    
    Call RefreshStateInfo
    
    '检查是否设置病区标记分类(-1不进行提示)
    If lngPreIdx <> -1 Then
        If mLngCount = 0 Then
            'MsgBox "病区【" & Split(mstr病区名称, "-")(1) & "】还未设置病区标记分类,请添加.", vbInformation, gstrSysName
        End If
    End If
    
    Screen.MousePointer = 0
    RefreshData = True
    Exit Function
ErrHand:
    Screen.MousePointer = 0
    If ErrCenter = 1 Then
        Resume
        Call SaveErrLog
    End If
End Function


Private Function UnitRefresh(Optional lngPreIdx As Long, Optional str标记 As String = "") As Boolean
'-----------------------------------------------
'功能:标记项目新增，修改后定位到选择的记录
'参数:lngreIdx 上次选择列的索引
'     str标记 上次选择列的内容 格式:主题序号-标记序号-病区ID
'-----------------------------------------------
    Dim objRow As ReportRow, i As Long, j As Long
    Dim blnRetrun As Boolean, blnChild As Boolean
    Dim arrCode() As String
    Dim lngRow As Long, lngGroup As Long
    
    If lngPreIdx < 0 Then lngPreIdx = 0
    
    If str标记 <> "" Then
        
        str标记 = str标记 & String(3 - UBound(Split(str标记, "-")), "-")
        arrCode = Split(str标记, "-")
        blnChild = Val(arrCode(1)) <> 0
        
        If blnChild = True Then
            If GetMarkCount = 0 Then blnChild = False
        End If
        
        If blnChild = True Then
            '先快速定位
            If lngPreIdx <= UnitReportControl.Rows.Count - 1 Then
                If Not UnitReportControl.Rows(lngPreIdx).GroupRow And UnitReportControl.Rows(lngPreIdx).Childs.Count = 0 Then
                    If UnitReportControl.Rows(lngPreIdx).Record(COL_主题序号).Record.Tag = str标记 Then
                        Set objRow = UnitReportControl.Rows(lngPreIdx)
                    End If
                End If
            End If
            '再进行查找
            If objRow Is Nothing Then
                For i = 0 To UnitReportControl.Rows.Count - 1
                    If Not UnitReportControl.Rows(i).GroupRow And UnitReportControl.Rows(i).Childs.Count = 0 Then
                        If UnitReportControl.Rows(i).Record(COL_主题序号).Record.Tag = str标记 Then
                            Set objRow = UnitReportControl.Rows(i): Exit For
                        End If
                    End If
                Next
            End If
        Else
            For i = 0 To UnitReportControl.Rows.Count - 1
                   If UnitReportControl.Rows(i).GroupRow And UnitReportControl.Rows(i).Childs.Count > 0 Then
                        If Split(UnitReportControl.Rows(i).Childs(0).Record(COL_主题序号).Record.Tag, "-")(0) = arrCode(0) And arrCode(1) = 0 Then
                            Set objRow = UnitReportControl.Rows(i): Exit For
                        End If
                   End If
            Next i
        End If
    End If
    
    '取第一个非分组行
    If objRow Is Nothing Then
        For i = 0 To UnitReportControl.Rows.Count - 1
            If blnChild Then
                If Not UnitReportControl.Rows(i).GroupRow And UnitReportControl.Rows(i).Childs.Count = 0 Then
                    If Val(Split(UnitReportControl.Rows(i).Record(COL_主题序号).Record.Tag, "-")(1)) <> 0 Then
                        Set objRow = UnitReportControl.Rows(i): Exit For
                    End If
                End If
            Else
                Set objRow = UnitReportControl.Rows(i)
                If objRow.GroupRow And objRow.Childs.Count > 0 Then
                    For j = 0 To objRow.Childs.Count - 1
                        If Val(Split(objRow.Childs(j).Record(COL_主题序号).Record.Tag, "-")(1)) <> 0 Then
                            Set objRow = UnitReportControl.Rows(i + 1)
                            Exit For
                        End If
                    Next j
                End If
                Exit For
            End If
        Next
    End If
    
    If Not objRow Is Nothing Then
        blnRetrun = True
        If Not objRow.GroupRow Then
            If Val(Split(objRow.Record(COL_主题序号).Record.Tag, "-")(1)) = 0 Then
                Set objRow = UnitReportControl.Rows(objRow.Index - 1)
            End If
        End If
        Set UnitReportControl.FocusedRow = objRow '该行选中且显示在可见区域,并引发SelectionChanged事件
        UnitReportControl.FocusedRow.Selected = True
        
    End If
    
    UnitRefresh = blnRetrun
End Function

Private Function AddRecord(ByVal strUnit As String, ByVal lngImgIndex As Long, ByVal StrInfo As String, ByVal lngDay As Long, _
    Optional ByVal strUnitOld As String = "") As ReportRecord
'-------------------------------------------------------------------------------------------
'功能：向ReportRecord添加病区标记记录
'------------------------------------------------------------------------------------------
    Dim blnParent As Boolean
    Dim Record As ReportRecord
    Set Record = UnitReportControl.Records.Add()
    
    If strUnitOld = "" Then strUnitOld = strUnit
    Dim Item As ReportRecordItem
   
    blnParent = Val(Split(strUnit, "-")(1)) = 0
    
    Set Item = Record.AddItem("")
    If blnParent Then Item.BackColor = RGB(255, 255, 255)
    
    Set Item = Record.AddItem("")
    If lngImgIndex >= mlngImgIndex And lngImgIndex <= zlCommFun.GetPaitSignImageList(0).ListImages.Count - 1 And blnParent = False Then
        Item.Icon = lngImgIndex
    End If
    If blnParent Then Item.BackColor = RGB(255, 255, 255)
    
    Set Item = Record.AddItem(StrInfo)
    If blnParent Then Item.BackColor = RGB(255, 255, 255)
    
    Set Item = Record.AddItem(Val(Split(strUnit, "-")(0)))
    Item.GroupCaption = "分组：" & Val(Split(strUnit, "-")(0)) & "-" & mstrSubject
    '主题序号 & "-" & 标记序号 & "-" & 病区Id & "-" & "有效天数"
    Item.Record.Tag = strUnit & "-" & lngDay
    
    Set Item = Record.AddItem(IIF(blnParent, "", lngDay)) '有效天数
    If blnParent Then Item.BackColor = RGB(255, 255, 255)
    Record.AddItem CInt(Split(strUnitOld, "-")(0))  '记录原始主题序号
    Record.AddItem CInt(Split(strUnitOld, "-")(1)) '记录原始标记序号
    Record.AddItem mstrSubject
    Record.AddItem mintSpecial
    
    Set AddRecord = Record
End Function

Private Function InitUnits() As Boolean
'功能：初始化住院护理病区
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String, i As Long
    Dim blnTrue As Boolean
    On Error GoTo errH
    
    '114577:支持设置公共分组图标
     strSql = _
         " Select Distinct A.ID,A.编码,A.名称" & _
         " From 部门表 A,部门性质说明 B " & _
         " Where A.ID=B.部门ID And B.服务对象 in(1,2,3) And B.工作性质='护理'" & _
         " And (A.撤档时间 is NULL or Trunc(A.撤档时间)=To_Date('3000-01-01','YYYY-MM-DD'))" & _
         " And (A.站点='" & gstrNodeNo & "' Or A.站点 is Null)" & _
         " Order by A.编码"

    cboUnit.Clear
    cboUnit.AddItem "0-公共病区"
    cboUnit.ItemData(cboUnit.NewIndex) = 0
    
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, glngUserId)
    
    If Not rsTmp.EOF Then
        For i = 1 To rsTmp.RecordCount
            cboUnit.AddItem rsTmp!编码 & "-" & rsTmp!名称
            cboUnit.ItemData(cboUnit.NewIndex) = rsTmp!ID
            
            If m病区ID = rsTmp!ID Then
                Call zlControl.CboSetIndex(cboUnit.hwnd, cboUnit.NewIndex)
                If cboUnit.ListIndex <> -1 Then blnTrue = True
            End If
            
            If Not blnTrue Then
                If rsTmp!ID = glngDeptId Then  '直接所属优先
                    Call zlControl.CboSetIndex(cboUnit.hwnd, cboUnit.NewIndex)
                End If
            End If
            rsTmp.MoveNext
        Next
    End If
    
    If cboUnit.ListIndex = -1 And cboUnit.ListCount > 0 Then
        Call zlControl.CboSetIndex(cboUnit.hwnd, 0)
    End If
    
    If cboUnit.ListIndex <> -1 Then
        m病区ID = cboUnit.ItemData(cboUnit.ListIndex)
        mstr病区名称 = cboUnit.Text
    End If
    
    InitUnits = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Resize()
    Call ResizeState
End Sub

Private Sub SetControlEnable(Optional blnEnable As Boolean = False)
'------------------------------------------------------------------
'功能:设置是否可以编辑
'------------------------------------------------------------------
        Dim blnNone As Boolean
        Dim i As Integer
        cbo标记.Enabled = blnEnable
       
        cbo标记.BackColor = IIF(blnEnable = False, Enable_Color, UnEnable_Color)
        
        blnNone = IIF(fraInfo.Tag = "新增", True, False)
        
        If blnNone = False Then
            If UnitReportControl.SelectedRows.Count > 0 Then
                If Not UnitReportControl.SelectedRows(0).GroupRow And UnitReportControl.SelectedRows(0).Childs.Count = 0 Then
                    blnNone = False
                Else
                    blnNone = True
                End If
            Else
                blnNone = True
            End If
        End If
        
        If UnitReportControl.Records.Count = 0 Then
            cbo标记.ListIndex = -1
        Else
            If UnitReportControl.SelectedRows.Count > 0 Then
                If Not UnitReportControl.SelectedRows(0).GroupRow And UnitReportControl.SelectedRows(0).Childs.Count = 0 Then
                    cbo标记.ListIndex = SetCboIndex(cbo标记, Val(Split(UnitReportControl.SelectedRows(0).Record(COL_主题序号).Record.Tag, "-")(0)))
                Else
                    cbo标记.ListIndex = SetCboIndex(cbo标记, Val(Split(UnitReportControl.SelectedRows(0).Childs(0).Record(COL_主题序号).Record.Tag, "-")(0)))
                End If
            End If
        End If
        
        If blnNone = True Then lblSet(9).Tag = "": cbo标记.Tag = ""
        txtInfo.Enabled = blnEnable
        txtInfo.BackColor = IIF(blnEnable = False, Enable_Color, UnEnable_Color)
        If blnNone Then txtInfo.Text = "": lblSet(8).Tag = "":: txtInfo.Tag = ""
        imaCustom.Enabled = blnEnable
        imaCustom.Locked = True
        imaCustom.BackColor = IIF(blnEnable = False, Enable_Color, UnEnable_Color)
        If blnNone Then imaCustom.Text = "": lblSet(7).Tag = "": imaCustom.Tag = ""
        
        cmdImage.Enabled = blnEnable
        
        If blnEnable = True And fraInfo.Visible = True Then cbo标记.SetFocus
End Sub

Private Sub ResizeState()
'功能:设置窗体所有控件位置
    Dim lngLeft As Long, lngTop As Long, lngRight As Long, lngBottom As Long
    Dim blnGourp As Boolean
    Dim objRow As ReportRow
    Dim i As Integer
    
    If Me.WindowState = 1 Then Exit Sub
    Call cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    If lngTop = 0 Then lngTop = 600
    
    mRect.Top = lngTop
    mRect.Left = lngLeft
    mRect.Right = lngRight
    mRect.Bottom = lngBottom
    
    fraUd.Top = lngTop
    fraUd.Left = 0
    fraUd.Width = ScaleWidth * 0.6
    fraUd.Height = lngBottom - lngTop
    
    UnitReportControl.Move 0, 100, fraUd.Width - 50, fraUd.Height - 150
    
    fraLine.Width = 50
    fraLine.Top = lngTop
    fraLine.Left = ScaleWidth * 0.6
    fraLine.Height = lngBottom - lngTop

    If InStr(1, ",新增,修改,", "," & fraInfo.Tag & ",") = 0 And InStr(1, ",新增,修改,", "," & fraUnit.Tag & ",") = 0 Then
        blnGourp = False
        If UnitReportControl.Rows.Count > 0 Then
            If GetMarkCount > 0 Then
                For i = 0 To UnitReportControl.Rows.Count - 1
                    If UnitReportControl.Rows(i).Selected = True Then
                        Set objRow = UnitReportControl.Rows(i)
                    End If
                Next i
                
                If Not objRow Is Nothing Then
                    If objRow.GroupRow Then
                        blnGourp = True
                    Else
                        blnGourp = False
                    End If
                Else
                    blnGourp = False
                End If
            Else
                blnGourp = True
            End If
        Else
            blnGourp = True
        End If
    ElseIf InStr(1, ",新增,修改,", "," & fraInfo.Tag & ",") = 0 Then
        blnGourp = True
    Else
        blnGourp = False
    End If
    
    Call SetFraResize(blnGourp)
End Sub

Private Sub SetFraResize(Optional blnGroup As Boolean = False)
    If blnGroup = True Then
        fraInfo.Visible = False
        fraInfo.Enabled = False
        fraUnit.Visible = True
        fraUnit.Enabled = True
        fraUnit.Top = mRect.Top
        fraUnit.Width = ScaleWidth * 0.4 - fraLine.Width
        fraUnit.Height = mRect.Bottom - mRect.Top
        fraUnit.Left = ScaleWidth * 0.6 + fraLine.Width
    Else
        fraUnit.Visible = False
        fraUnit.Enabled = False
        fraInfo.Visible = True
        fraInfo.Enabled = True
        fraInfo.Top = mRect.Top
        fraInfo.Width = ScaleWidth * 0.4 - fraLine.Width
        fraInfo.Height = mRect.Bottom - mRect.Top
        fraInfo.Left = ScaleWidth * 0.6 + fraLine.Width
    End If
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If stbThis.Visible Then Bottom = stbThis.Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstrSubject = ""
    mlngDay = 0
    mintSpecial = 0
    Call UnLoadImage
    mblnOK = (fraUd.Tag = "1")
    If Not (mrsData Is Nothing) Then Set mrsData = Nothing
'    If mblnChange = True Then
'        If MsgBox("病区【" & Split(mstr病区名称, "-")(1) & "】标记内容已经发生改变，你确定要退出吗?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Cancel = 1
'    End If
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub HScr_Change()
    pic标记.Top = HScr.Top - HScr.Value
    If picBack.Visible = True Then picBack.SetFocus
End Sub

Private Sub HScr_Scroll()
    pic标记.Top = HScr.Top - HScr.Value
End Sub

Private Sub imaCustom_Click()
     Call showIcon(imaCustom.SelectedItem.Index - 1)
End Sub

Private Sub imaCustom_GotFocus()
    If picBack.Visible = True Then
        picBack.Visible = False
        cmdImage.Enabled = True
    End If
End Sub

Private Sub imaCustom_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    If KeyAscii <> vbKeyReturn Then
        Call zlControl.CboMatchIndex(imaCustom.hwnd, KeyAscii)
    Else
        '由于敲回车后ImageCombo图形丢失，此处重新显示图标
        If KeyAscii = vbKeyReturn Then
            If imaCustom.Text <> "" Then
                 For i = 1 To zlCommFun.GetPaitSignImageList(0).ListImages.Count - mlngImgIndex
                    If imaCustom.Text = zlCommFun.GetPaitSignImageList(0).ListImages(mlngImgIndex + i).Key Then
                        imaCustom.ComboItems(i).Selected = True
                    End If
                Next i
            End If
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub

Private Sub imgIcon_DblClick(Index As Integer)
    Call showIcon(Index)
End Sub

Private Sub showIcon(ByVal Index As Integer)
'功能:展示用户选择的图标
    If Index < 0 Then Exit Sub
    imaCustom.ComboItems(Index + 1).Selected = True
    picBack.Visible = False
    cmdImage.Enabled = True
    
    If fraInfo.Tag = "修改" Then
        With UnitReportControl.FocusedRow.Record(COL_标注)
            .Icon = Index + mlngImgIndex
        End With
        UnitReportControl.Populate
    End If
    
    If (txtInfo.Text = "" Or txtInfo.Tag <> "改变") And IIF(fraInfo.Tag = "修改", lblSet(8).Tag = "", True) Then txtInfo.Text = imaCustom.ComboItems(Index + 1).Text
End Sub

Private Sub ShowSelect(ByVal Index As Integer)
'功能:选中图标
    Dim i As Integer
    lblSelect(Index).BackColor = &H8000000D
    lblInfo(Index).BackColor = &H8000000D
    For i = 0 To zlCommFun.GetPaitSignImageList(1).ListImages.Count - mlngImgIndex - 1
        If i <> Index Then
            lblSelect(i).BackColor = &H8000000E
            lblInfo(i).BackColor = &H8000000E
        End If
    Next i
End Sub

Private Sub imgIcon_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call ShowSelect(Index)
End Sub

Private Function AgainComputePreId(ByVal lngPreVId As Long, Optional bln新增 As Boolean = False) As Long
'--------------------------------------
'功能:计算算标记序号
'参数：lngPreVId：主题序号
'--------------------------------------
    Dim lngTmp As Long
    Dim blnTrue As Boolean
    Dim i As Integer
    For i = 0 To UnitReportControl.Records.Count - 1
        If lngPreVId = Val(Split(UnitReportControl.Records(i).Item(COL_主题序号).Record.Tag, "-")(0)) Then
            If lngTmp < Val(Split(UnitReportControl.Records(i).Item(COL_主题序号).Record.Tag, "-")(0)) Then
                lngTmp = Val(Split(UnitReportControl.Records(i).Item(COL_主题序号).Record.Tag, "-")(0))
            End If
        End If
    Next i
    
    If bln新增 = True Then
        '新增的记录直接加一
        lngTmp = lngTmp + 1
    Else
        '个性标记改变时如果和以前不同就序号直接加一，如果回复到以前则检测以前序号是否被使用，使用的话重新获取新的序号
        If Val(Split(UnitReportControl.FocusedRow.Record(COL_主题序号).Record.Tag, "-")(0)) = lngPreVId Then
            '检查原始序号是否被新增记录使用
            For i = 0 To UnitReportControl.Records.Count - 1
                If lngPreVId = Val(Split(UnitReportControl.Records(i).Item(COL_主题序号).Record.Tag, "-")(0)) Then
                    If UnitReportControl.FocusedRow.Record(COL_原始标记).Value = Val(Split(UnitReportControl.Records(i).Item(COL_主题序号).Record.Tag, "-")(1)) Then
                        blnTrue = True
                    End If
                End If
            Next i
            
            If blnTrue = True Then
                lngTmp = UnitReportControl.FocusedRow.Record(COL_原始标记).Value
            Else
                lngTmp = lngTmp + 1
            End If
        Else
            lngTmp = lngTmp + 1
        End If
    End If

    AgainComputePreId = lngTmp
    
End Function


Private Function SaveData() As Boolean
'------------------------------------------------------------------
'功能：病区标记数据保存
'------------------------------------------------------------------
    Dim lngRowIndex As Long '选择列的索引
    Dim i As Integer
    Dim Record As ReportRecord
    Dim strTemp As String, strSql As String
    Dim blnTran As Boolean
    Dim strSQLAdd() As String
    Dim StrSQLMod() As String
    Dim strTmp1 As String
    Dim rsTmp As New ADODB.Recordset
    On Error GoTo ErrHand
    
    ReDim Preserve strSQLAdd(0 To 0)
    ReDim Preserve StrSQLMod(0 To 0)
    lngRowIndex = 0
    
    If InStr(1, ",新增,修改,", "," & fraInfo.Tag & ",") <> 0 Then
        If imaCustom.Text = "" Then
            MsgBox "标记图形不能为空,请选择标记图形后在进行保存操作.", vbInformation, gstrSysName
            imaCustom.SetFocus
            Exit Function
        End If
    End If
    
    If InStr(1, ",新增,修改,", "," & fraUnit.Tag & ",") <> 0 Then
        If Trim(txtName.Text) = "" Then
            MsgBox "标记名称不能为空,请检查.", vbInformation, gstrSysName
            txtName.SetFocus
            Exit Function
        End If
        
        If Not zlCommFun.StrIsValid(txtDays.Text, 3, txtDays.hwnd, "有效天数") Then Exit Function
    End If
    
    '修改
    If fraInfo.Tag = "修改" Then
        If UnitReportControl.FocusedRow Is Nothing Then Exit Function
        
        lngRowIndex = UnitReportControl.FocusedRow.Index
        mUnit.病区ID = m病区ID
        mUnit.主题序号 = Val(Split(UnitReportControl.Rows(lngRowIndex).Record(COL_主题序号).Record.Tag, "-")(0))
        mUnit.标记序号 = Val(Split(UnitReportControl.Rows(lngRowIndex).Record(COL_主题序号).Record.Tag, "-")(1))
        mUnit.说明 = zlCommFun.Nvl(UnitReportControl.Rows(lngRowIndex).Record(COL_说明).Value)
        mUnit.说明 = Trim(txtInfo.Text)
        mUnit.图形索引 = Val(zlCommFun.Nvl(UnitReportControl.Rows(lngRowIndex).Record(COL_标注).Icon, 0))
        mUnit.有效天数 = Val(zlCommFun.Nvl(UnitReportControl.Rows(lngRowIndex).Record(COL_有效天数).Value, 0))
        mUnit.原始主题 = zlCommFun.Nvl(UnitReportControl.Rows(lngRowIndex).Record(COL_原始主题).Value, 0)
        mUnit.原始标记 = zlCommFun.Nvl(UnitReportControl.Rows(lngRowIndex).Record(COL_原始标记).Value, 0)
        
        mrsData.Filter = "主题序号=" & Val(mUnit.主题序号) & " and 标记序号=0"
        If mrsData.RecordCount > 0 Then
            mUnit.有效天数 = Val(Nvl(mrsData!有效天数))
        End If
        
        '修改后数据无任何变化,不进行数据写入操作
        If CheckChange Then
            If mUnit.主题序号 <> mUnit.原始主题 Then '主题序号发生改变
                StrSQLMod(ReDimArray(StrSQLMod)) = "Zl_病区标记内容_Delete(" & mUnit.病区ID & "," & mUnit.原始主题 & "," & mUnit.原始标记 & ")"
                mUnit.标记序号 = GetNewPreID(mUnit.病区ID, mUnit.主题序号)
                
                strTmp1 = mUnit.主题序号 & "-" & mUnit.标记序号
                StrSQLMod(ReDimArray(StrSQLMod)) = "Zl_病区标记内容_Insert(" & mUnit.病区ID & "," & mUnit.主题序号 & "," & _
                mUnit.标记序号 & ",'" & mUnit.说明 & "'," & mUnit.图形索引 & "," & mUnit.有效天数 & ")"
            Else
                strTmp1 = mUnit.主题序号 & "-" & mUnit.原始标记
                StrSQLMod(ReDimArray(StrSQLMod)) = "Zl_病区标记内容_Update(" & mUnit.病区ID & "," & mUnit.主题序号 & "," & _
                    mUnit.原始标记 & ",'" & mUnit.说明 & "'," & mUnit.图形索引 & "," & mUnit.有效天数 & ")"
            End If
            
            If IsEqualInfo(txtInfo.Text, False, strTmp1) = False Then
                If txtInfo.Enabled And txtInfo.Visible Then txtInfo.SetFocus
                Exit Function
            End If
                
            If UBound(StrSQLMod) > 1 Then
                gcnOracle.BeginTrans
                blnTran = True
                For i = 0 To UBound(StrSQLMod)
                    If StrSQLMod(i) <> "" Then Call zlDatabase.ExecuteProcedure(StrSQLMod(i), Me.Caption)
                Next i
                gcnOracle.CommitTrans
            Else
                For i = 0 To UBound(StrSQLMod)
                    If StrSQLMod(i) <> "" Then Call zlDatabase.ExecuteProcedure(StrSQLMod(i), Me.Caption)
                Next i
            End If
            
            fraUd.Tag = "1"
        Else
            strTmp1 = mUnit.主题序号 & "-" & mUnit.原始标记
        End If
        strTemp = strTmp1 & "-" & mUnit.病区ID & "-" & Val(mUnit.有效天数)
    End If
    
    '新增
    If fraInfo.Tag = "新增" Then
        If cbo标记.ListIndex = -1 Then Exit Function
        If IsEqualInfo(txtInfo.Text, False) = False Then
            If txtInfo.Enabled And txtInfo.Visible Then txtInfo.SetFocus
            Exit Function
        End If
        mUnit.病区ID = m病区ID
        mUnit.主题序号 = cbo标记.ItemData(cbo标记.ListIndex)
        mUnit.标记序号 = GetNewPreID(mUnit.病区ID, mUnit.主题序号)
        mUnit.说明 = txtInfo.Text
        mUnit.图形索引 = imaCustom.SelectedItem.Index - 1 + mlngImgIndex
        mUnit.有效天数 = 0
        
        For i = 0 To UnitReportControl.Rows.Count - 1
            If Not UnitReportControl.Rows(i).GroupRow And UnitReportControl.Rows(i).Childs.Count = 0 Then
                If Val(Split(UnitReportControl.Rows(i).Record(COL_主题序号).Record.Tag, "-")(0)) = cbo标记.ItemData(cbo标记.ListIndex) Then
                    mUnit.有效天数 = Val(Split(UnitReportControl.Rows(i).Record(COL_有效天数).Record.Tag, "-")(3))
                    Exit For
                End If
            End If
        Next i
        
        mrsData.Filter = "主题序号=" & Val(mUnit.主题序号) & " and 标记序号=0"
        If mrsData.RecordCount > 0 Then
            mUnit.有效天数 = Val(Nvl(mrsData!有效天数))
        End If
        
        strTmp1 = mUnit.主题序号 & "-" & mUnit.标记序号
        
        strSQLAdd(ReDimArray(strSQLAdd)) = "Zl_病区标记内容_Insert(" & mUnit.病区ID & "," & mUnit.主题序号 & "," & _
            mUnit.标记序号 & ",'" & mUnit.说明 & "'," & mUnit.图形索引 & "," & mUnit.有效天数 & ")"
            
        For i = 0 To UBound(strSQLAdd)
            If strSQLAdd(i) <> "" Then Call zlDatabase.ExecuteProcedure(strSQLAdd(i), Me.Caption)
        Next i
        
        strTemp = strTmp1 & "-" & mUnit.病区ID & "-" & Val(mUnit.有效天数)
        
        mstrSubject = cbo标记.Text
        Set Record = AddRecord(mUnit.主题序号 & "-" & mUnit.标记序号 & "-" & mUnit.病区ID, mUnit.图形索引, mUnit.说明, Val(mUnit.有效天数))
        fraUd.Tag = "1"
        UnitReportControl.Populate
    End If
                
    '新增主题名称
    If fraUnit.Tag = "新增" Then
        If IsEqualInfo(txtName.Text, True) = False Then
            If txtName.Enabled And txtName.Visible Then txtName.SetFocus
            Exit Function
        End If
        mUnit.主题序号 = GetNewSubjectId(cboUnit.ItemData(cboUnit.ListIndex))
        If mUnit.主题序号 = 0 Then Exit Function
        
        strSQLAdd(ReDimArray(strSQLAdd)) = "Zl_病区标记内容_Insert(" & cboUnit.ItemData(cboUnit.ListIndex) & "," & mUnit.主题序号 & "," & _
            0 & ",'" & Replace(Trim(txtName.Text), "'", "") & "'," & 0 & "," & Val(txtDays.Text) & "," & IIF(chkSpecial.Value = 0, "NULL", "1") & ")"
        
        For i = 0 To UBound(strSQLAdd)
            If strSQLAdd(i) <> "" Then Call zlDatabase.ExecuteProcedure(strSQLAdd(i), Me.Caption)
        Next i
        
        strTemp = mUnit.主题序号 & "-0-" & mUnit.病区ID & "-" & Val(txtDays.Text)
        
        fraUd.Tag = "1"
    End If
    
    '修改主题名称
    If fraUnit.Tag = "修改" Then
        If UnitReportControl.Rows(UnitReportControl.Tag) Is Nothing Then Exit Function
        
        mUnit.主题序号 = Val(Split(UnitReportControl.Rows(UnitReportControl.Tag).Childs(0).Record(COL_主题序号).Record.Tag, "-")(0))
        mUnit.病区ID = cboUnit.ItemData(cboUnit.ListIndex)
        
        '标记分类发生变化则进行修改操作
        If CheckChange Then
            If IsEqualInfo(txtName.Text, True, mUnit.主题序号) = False Then
                If txtName.Enabled And txtName.Visible Then txtName.SetFocus
                Exit Function
            End If
            StrSQLMod(ReDimArray(StrSQLMod)) = "Zl_病区标记内容_Update(" & mUnit.病区ID & "," & mUnit.主题序号 & "," & _
                0 & ",'" & Replace(Trim(txtName.Text), "'", "") & "'," & 0 & "," & Val(txtDays.Text) & "," & IIF(chkSpecial.Value = 0, "NULL", "1") & ")"
            
            strSql = "select 标记序号,说明,图形索引,有效天数 from 病区标记内容 where " & IIF(mUnit.病区ID = 0, " 病区ID IS NULL ", " 病区ID=[1] ") & " and  主题序号=[2] and 标记序号<>0"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "病区标记内容", mUnit.病区ID, mUnit.主题序号)
            '检查子分类的天数是否和分类相同，不同则进行修改
            With rsTmp
                Do While Not .EOF
                    If zlCommFun.Nvl(!有效天数, 0) <> Val(txtDays.Text) Then
                        StrSQLMod(ReDimArray(StrSQLMod)) = "Zl_病区标记内容_Update(" & mUnit.病区ID & "," & mUnit.主题序号 & "," & _
                            zlCommFun.Nvl(!标记序号, 0) & ",'" & Replace(zlCommFun.Nvl(!说明), "'", "") & "'," & zlCommFun.Nvl(!图形索引, 0) & "," & Val(txtDays.Text) & ")"
                    End If
                .MoveNext
                Loop
            End With
            
            If UBound(StrSQLMod) > 1 Then
                gcnOracle.BeginTrans
                blnTran = True
                For i = 0 To UBound(StrSQLMod)
                    If StrSQLMod(i) <> "" Then Call zlDatabase.ExecuteProcedure(StrSQLMod(i), Me.Caption)
                Next i
                gcnOracle.CommitTrans
            Else
                For i = 0 To UBound(StrSQLMod)
                    If StrSQLMod(i) <> "" Then Call zlDatabase.ExecuteProcedure(StrSQLMod(i), Me.Caption)
                Next i
            End If
            fraUd.Tag = "1"
        End If
        strTemp = mUnit.主题序号 & "-0-" & mUnit.病区ID & "-" & Val(txtDays.Text)
    End If
    
    mblnChange = False
    
    fraInfo.Tag = ""
    fraUnit.Tag = ""
    UnitReportControl.Tag = ""
    '定位相应的列上
    Call RefreshData(lngRowIndex, strTemp)
    fraUd.Enabled = True
    UnitReportControl.SetFocus
    
    SaveData = True
    Exit Function
ErrHand:
    If blnTran = True Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngRowIndex As Long '选择列的索引
    Dim i As Integer
    Dim Record As ReportRecord
    Dim strTemp As String, strSql As String
    Dim blnTran As Boolean
    Dim cbrControl As CommandBarControl
    Dim strTmp1 As String
    Dim rsTmp As New ADODB.Recordset
    On Error GoTo ErrHand

    
    Select Case Control.ID
        Case conMenu_File_PrintSet
            Call zlPrintSet
                    
        Case conMenu_File_Preview
            Call zlRptPrint(2)
        
        Case conMenu_File_Print
            Call zlRptPrint(1)
        
        Case conMenu_File_Excel
            Call zlRptPrint(3)
    
        Case conMenu_View_ToolBar_Button
            cbsMain(2).Visible = Not cbsMain(2).Visible
            cbsMain.RecalcLayout
        
        Case conMenu_View_ToolBar_Text
            For Each cbrControl In cbsMain(2).Controls
                If cbrControl.Type <> xtpControlLabel Then
                    cbrControl.Style = IIF(cbrControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
                End If
            Next
            cbsMain.RecalcLayout
            
        Case conMenu_View_StatusBar
            stbThis.Visible = Not stbThis.Visible
            cbsMain.RecalcLayout
            
        Case conMenu_Edit_NewItem     '*新增
            fraInfo.Tag = "新增"
            fraUnit.Tag = ""
            Call SetFraResize
            Call SetControlEnable(True)
            mblnChange = True
        Case conMenu_Edit_Modify      '*修改(&M)
            fraInfo.Tag = "修改"
            fraUnit.Tag = ""
            Call SetControlEnable(True)
            mblnChange = True
            
        Case conMenu_Edit_Delete      '*删除(&D)
            If MsgBox("你确定要删除病区【" & Split(mstr病区名称, "-")(1) & "】内容【" & UnitReportControl.FocusedRow.Record(COL_说明).Value & "】的标记信息吗?", _
                vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            
            strTemp = UnitReportControl.FocusedRow.Record(COL_主题序号).Record.Tag
            
            mUnit.病区ID = CInt(Split(strTemp, "-")(2))
            mUnit.主题序号 = CInt(Split(strTemp, "-")(0))
            mUnit.标记序号 = CInt(Split(strTemp, "-")(1))
            
            '检查改主题内容该病区是否正在使用
            If CheckUseUnit(mUnit.病区ID, mUnit.主题序号, mUnit.标记序号) = True Then Exit Sub
            
            strSql = "Zl_病区标记内容_Delete(" & mUnit.病区ID & "," & mUnit.主题序号 & "," & mUnit.标记序号 & ")"
            Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
            
            '定位到下一列
            lngRowIndex = UnitReportControl.FocusedRow.Index
            
            Call UnitReportControl.Records.RemoveAt(UnitReportControl.FocusedRow.Record.Index)
            UnitReportControl.Populate
            
            If UnitReportControl.Records.Count > 0 Then
                lngRowIndex = IIF(UnitReportControl.Rows.Count - 1 > lngRowIndex, lngRowIndex, UnitReportControl.Rows.Count - 1)
                
                If UnitReportControl.Rows(lngRowIndex).GroupRow And UnitReportControl.Rows(lngRowIndex).Childs.Count <> 0 Then
                    lngRowIndex = lngRowIndex - 1
                End If
                
                If UnitReportControl.Rows(lngRowIndex).GroupRow Then
                    strTemp = UnitReportControl.Rows(lngRowIndex).Childs.Record(COL_主题序号).Record.Tag
                Else
                    strTemp = UnitReportControl.Rows(lngRowIndex).Record(COL_主题序号).Record.Tag
                End If
            End If
            Call RefreshData(lngRowIndex, strTemp)
            mblnChange = False
            fraUd.Tag = "1"
            fraUd.Enabled = True
            UnitReportControl.SetFocus
        Case conMenu_Edit_NewParent '*新增分类
            fraInfo.Tag = ""
            fraUnit.Tag = "新增"
            Call SetFraResize(True)
            txtName.Enabled = True
            txtName.Text = ""
            txtDays.Enabled = True
            txtDays.Text = ""
            txtName.BackColor = UnEnable_Color
            txtDays.BackColor = UnEnable_Color
            chkSpecial.Visible = (m病区ID = 0)
            chkSpecial.Enabled = (m病区ID = 0)
            chkSpecial.Value = 0
            If m病区ID = 0 Then
                mrsData.Filter = "标记序号=0"
                Do While Not mrsData.EOF
                    If Val("" & mrsData!是否特殊) = 1 Then
                        chkSpecial.Enabled = False
                        Exit Do
                    End If
                    mrsData.MoveNext
                Loop
            End If
            txtName.SetFocus
            UnitReportControl.Tag = ""
            mblnChange = True
            
        Case conMenu_Edit_ModifyParent ' "修改分类(&U)"
            fraInfo.Tag = ""
            fraUnit.Tag = "修改"
            txtName.Enabled = True
            txtDays.Enabled = True
            chkSpecial.Visible = (m病区ID = 0)
            chkSpecial.Enabled = (m病区ID = 0)
            If m病区ID = 0 Then
                mrsData.Filter = "标记序号=0 and 主题序号<>" & Val(Split(UnitReportControl.FocusedRow.Childs(0).Record(COL_主题序号).Record.Tag, "-")(0))
                Do While Not mrsData.EOF
                    If Val("" & mrsData!是否特殊) = 1 Then
                        chkSpecial.Enabled = False
                        Exit Do
                    End If
                    mrsData.MoveNext
                Loop
            End If
            txtName.BackColor = UnEnable_Color
            txtDays.BackColor = UnEnable_Color
            txtName.SetFocus
            UnitReportControl.Tag = UnitReportControl.FocusedRow.Index
            mblnChange = True

        Case conMenu_Edit_DeleteParent '"删除分类(&E)"
            If UnitReportControl.FocusedRow Is Nothing Then Exit Sub
            
            If MsgBox("你确定要删除病区【" & Split(mstr病区名称, "-")(1) & "】标记分类【" & UnitReportControl.FocusedRow.Childs(0).Record(COL_主题序号).GroupCaption & "】的信息吗?", _
                vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            
            strTemp = UnitReportControl.FocusedRow.Childs(0).Record(COL_主题序号).Record.Tag
            
            mUnit.病区ID = CInt(Split(strTemp, "-")(2))
            mUnit.主题序号 = CInt(Split(strTemp, "-")(0))
            mUnit.标记序号 = 0
            
            '检查改主题内容该病区是否正在使用
            If CheckUseUnit(mUnit.病区ID, mUnit.主题序号, mUnit.标记序号) = True Then Exit Sub
            
            strSql = "Zl_病区标记内容_Delete(" & mUnit.病区ID & "," & mUnit.主题序号 & "," & mUnit.标记序号 & ")"
            Call zlDatabase.ExecuteProcedure(strSql, Me.Caption)
            
            Call RefreshData(-1)
            
            mblnChange = False
            fraUd.Tag = "1"
            fraUd.Enabled = True
            UnitReportControl.SetFocus
            
        Case conMenu_Edit_Save     '*保存
            picBack.Visible = False
            cmdImage.Enabled = True
            Call SaveData
        Case conMenu_Edit_Reuse    '*取消
            '记录现在选中的标注
            If UnitReportControl.SelectedRows.Count > 0 Then
                If Not UnitReportControl.SelectedRows(0) Is Nothing Then
                    If Not UnitReportControl.SelectedRows(0).GroupRow And UnitReportControl.SelectedRows(0).Childs.Count = 0 Then
                        lngRowIndex = UnitReportControl.SelectedRows(0).Index '用于快速重新定位
                        strTemp = UnitReportControl.SelectedRows(0).Record(COL_主题序号).Record.Tag
                    Else
                        lngRowIndex = UnitReportControl.SelectedRows(0).Index '用于快速重新定位
                        strTmp1 = UnitReportControl.SelectedRows(0).Childs(0).Record(COL_主题序号).Record.Tag
                        strTemp = Split(strTmp1, "-")(0) & "-0-" & Split(strTmp1, "-")(2) & "-" & Split(strTmp1, "-")(3)
                    End If
                End If
            Else
                If UnitReportControl.Tag <> "" Then
                    If Not UnitReportControl.Rows(UnitReportControl.Tag) Is Nothing Then
                        If Not UnitReportControl.Rows(UnitReportControl.Tag).GroupRow And UnitReportControl.Rows(UnitReportControl.Tag).Childs.Count = 0 Then
                            lngRowIndex = UnitReportControl.Rows(UnitReportControl.Tag).Index
                            strTemp = UnitReportControl.Rows(UnitReportControl.Tag).Record(COL_主题序号).Record.Tag
                        Else
                            lngRowIndex = UnitReportControl.Rows(UnitReportControl.Tag).Index
                            strTmp1 = UnitReportControl.Rows(UnitReportControl.Tag).Childs(0).Record(COL_主题序号).Record.Tag
                            strTemp = Split(strTmp1, "-")(0) & "-0-" & Split(strTmp1, "-")(2) & "-" & Split(strTmp1, "-")(3)
                        End If
                    End If
                End If
            End If
            picBack.Visible = False
            cmdImage.Enabled = True
            fraInfo.Tag = ""
            fraUnit.Tag = ""
            Call RefreshData(lngRowIndex, strTemp)
            mblnChange = False
            fraUd.Enabled = True
            UnitReportControl.SetFocus
            
        Case conMenu_View_Refresh  '刷新
            '记录现在选中的标注
            If UnitReportControl.SelectedRows.Count > 0 Then
                If Not UnitReportControl.SelectedRows(0) Is Nothing Then
                    If Not UnitReportControl.SelectedRows(0).GroupRow And UnitReportControl.SelectedRows(0).Childs.Count = 0 Then
                        lngRowIndex = UnitReportControl.SelectedRows(0).Index '用于快速重新定位
                        strTemp = UnitReportControl.SelectedRows(0).Record(COL_主题序号).Record.Tag
                    Else
                        lngRowIndex = UnitReportControl.SelectedRows(0).Index '用于快速重新定位
                        strTmp1 = UnitReportControl.SelectedRows(0).Childs(0).Record(COL_主题序号).Record.Tag
                        strTemp = Split(strTmp1, "-")(0) & "-0-" & Split(strTmp1, "-")(2) & "-" & Split(strTmp1, "-")(3)
                    End If
                End If
            End If
            
            fraInfo.Tag = ""
            fraUnit.Tag = ""
            Call RefreshData(lngRowIndex, strTemp)
            mblnChange = False
            fraUd.Enabled = True
            UnitReportControl.SetFocus
            
        Case conMenu_Help_About
            Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
            
        Case conMenu_Help_Web_Home
            Call zlHomePage(Me.hwnd)
            
        Case conMenu_Help_Web_Forum '中联论坛
            Call zlWebForum(Me.hwnd)

        Case conMenu_Help_Web_Mail '发送Email
            Call zlMailTo(Me.hwnd)
            
        Case conMenu_Help_Help        '*帮助主题(&H)
             Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
        Case conMenu_File_Exit        '*退出(&X)
            Unload Me
    End Select
    
    Call RefreshStateInfo
    cbsMain.RecalcLayout
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveData
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case conMenu_File_Preview, conMenu_File_Print, conMenu_File_Excel
            Control.Enabled = (UnitReportControl.Records.Count > 0)
        Case conMenu_Edit_NewItem   '*新增(&A)
            If UnitReportControl.Rows.Count > 0 Then
                Control.Enabled = Not UnitReportControl.FocusedRow Is Nothing
                If Control.Enabled = True Then
                    Control.Enabled = Not mblnChange
                End If
            Else
                Control.Enabled = mLngCount > 0
            End If
        Case conMenu_Edit_Modify      '*修改(&M)
            If UnitReportControl.Rows.Count > 0 Then
                Control.Enabled = Not UnitReportControl.FocusedRow Is Nothing
                If Control.Enabled = True Then Control.Enabled = Not UnitReportControl.FocusedRow.GroupRow
                If Control.Enabled = True Then
                    Control.Enabled = Not mblnChange And Val(Split(UnitReportControl.FocusedRow.Record(COL_主题序号).Record.Tag, "-")(1)) <> 0
                End If
            Else
                Control.Enabled = False
            End If
        Case conMenu_Edit_Delete      '*删除(&D)
            If UnitReportControl.Rows.Count > 0 Then
                Control.Enabled = Not UnitReportControl.FocusedRow Is Nothing
                If Control.Enabled = True Then Control.Enabled = Not UnitReportControl.FocusedRow.GroupRow
                If Control.Enabled = True Then
                    Control.Enabled = Not mblnChange And Val(Split(UnitReportControl.FocusedRow.Record(COL_主题序号).Record.Tag, "-")(1)) <> 0
                End If
            Else
                Control.Enabled = False
            End If
        
        Case conMenu_Edit_NewParent '*新增分类
            Control.Enabled = Not UnitReportControl.FocusedRow Is Nothing
            If Control.Enabled = True Then
                Control.Enabled = Not mblnChange And UnitReportControl.FocusedRow.GroupRow
            Else
                If UnitReportControl.Rows.Count > 0 Then
                    Control.Enabled = Not mblnChange
                Else
                    Control.Enabled = True And Not mblnChange
                End If
            End If
             
        Case conMenu_Edit_ModifyParent ' "修改分类(&U)"
             If UnitReportControl.Rows.Count > 0 Then
                Control.Enabled = Not UnitReportControl.FocusedRow Is Nothing
                If Control.Enabled = True Then
                    Control.Enabled = Not mblnChange And UnitReportControl.FocusedRow.GroupRow
                End If
             Else
                Control.Enabled = False
             End If
        Case conMenu_Edit_DeleteParent '"删除分类(&E)"
             If UnitReportControl.Rows.Count > 0 Then
                Control.Enabled = Not UnitReportControl.FocusedRow Is Nothing
                If Control.Enabled = True Then
                    Control.Enabled = Not mblnChange And UnitReportControl.FocusedRow.GroupRow
                End If
             Else
                Control.Enabled = False
             End If
        Case conMenu_Edit_Save     '*保存
            Control.Enabled = mblnChange
        Case conMenu_Edit_Reuse     '*取消
            Control.Enabled = mblnChange
        Case conMenu_View_Refresh '*刷新
            Control.Enabled = Not mblnChange
        Case conMenu_View_ToolBar_Button
            Control.Checked = Me.cbsMain(2).Visible
        Case conMenu_View_ToolBar_Text
            Control.Checked = Not (Me.cbsMain(2).Controls(1).Style = xtpButtonIcon)
        Case conMenu_View_ToolBar_Size
            Control.Checked = Me.cbsMain.Options.LargeIcons
        Case conMenu_View_StatusBar
            Control.Checked = Me.stbThis.Visible
    End Select
    
    cboUnit.Enabled = Not mblnChange
    fraUd.Enabled = Not mblnChange
    
End Sub

Private Sub lblSelect_DblClick(Index As Integer)
    Call showIcon(Index)
End Sub

Private Sub lblSelect_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Call ShowSelect(Index)
End Sub

Private Sub picBack_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        picBack.Visible = False
        cmdImage.Enabled = True
    End If
End Sub

Private Sub picIcon_KeyPress(Index As Integer, KeyAscii As Integer)
    picBack_KeyPress KeyAscii
End Sub

Private Sub pic标记_KeyPress(KeyAscii As Integer)
    picBack_KeyPress KeyAscii
End Sub

Private Sub txtDays_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey (vbKeyTab)
    Else
        If KeyAscii > 45 And KeyAscii < 58 Then
            If KeyAscii = 46 Then
                If Len(txtDays.Text) = 0 Then
                    KeyAscii = 0
                Else
                    If InStr(1, txtDays.Text, ".") <> 0 Then
                        KeyAscii = 0
                    End If
                End If
            End If
        Else
            If KeyAscii <> 8 Then
                KeyAscii = 0
            End If
        End If
    End If
End Sub

Private Sub txtInfo_GotFocus()
    If picBack.Visible = True Then
        picBack.Visible = False
        cmdImage.Enabled = True
    End If
    txtInfo.SelStart = Len(txtInfo.Text)
    Call zlControl.TxtSelAll(txtInfo)
End Sub


Private Sub txtInfo_Change()
    If mblnChange = False Then Exit Sub
    
    If fraInfo.Tag = "修改" Then
        With UnitReportControl.FocusedRow.Record(COL_说明)
            .Value = txtInfo.Text
        End With
        UnitReportControl.Populate
    End If
    
    '判定操作员是否手工录入修改了标注说明
    If lblSet(8).Tag <> "" And lblSet(8).Tag <> Trim(txtInfo.Text) And Trim(txtInfo.Text) <> cmdImage.Tag Then
        txtInfo.Tag = "改变"
    End If
    
    If imaCustom.ComboItems.Count > 0 Then cmdImage.Tag = imaCustom.Text
End Sub

Private Sub txtInfo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If Trim(txtInfo.Text) <> "" Then
            txtInfo.Tag = "改变"
        End If
    Else
        If Chr(KeyCode) = "'" Or Chr(KeyCode) = "|" Then KeyCode = 0
    End If
End Sub


Private Sub txtName_Change()
    Dim i As Integer
    Dim lngPreIdx As Long
    Dim strTemp As String, str标记 As String
    If mblnChange = False Then Exit Sub
    
    If fraUnit.Tag = "修改" And UnitReportControl.Tag <> "" Then
        If UnitReportControl.Rows(UnitReportControl.Tag) Is Nothing Then Exit Sub
        With UnitReportControl.Rows(UnitReportControl.Tag)
            lngPreIdx = .Index
            strTemp = .Childs(0).Record(COL_主题序号).Record.Tag
            str标记 = Split(strTemp, "-")(0) & "-0-" & Split(strTemp, "-")(2) & "-" & Split(strTemp, "-")(3)
            
            For i = 0 To .Childs.Count - 1
                .Childs(i).Record(COL_主题序号).GroupCaption = "分组：" & Split(strTemp, "-")(0) & "-" & Replace(txtName.Text, "'", "")
            Next i
        End With
        UnitReportControl.Populate
    End If
End Sub

Private Sub txtName_GotFocus()
    txtName.SelStart = Len(txtName.Text)
    Call zlControl.TxtSelAll(txtName)
End Sub

Private Sub txtDays_GotFocus()
    txtDays.SelStart = Len(txtDays.Text)
    Call zlControl.TxtSelAll(txtDays)
End Sub

Private Sub txtDays_Change()
    Dim i As Integer
    If mblnChange = False Then Exit Sub
    '更改分类天数时，子分类同步更新
    If fraUnit.Tag = "修改" And UnitReportControl.Tag <> "" Then
        If UnitReportControl.Rows(UnitReportControl.Tag) Is Nothing Then Exit Sub
        With UnitReportControl.Rows(UnitReportControl.Tag)
            For i = 0 To .Childs.Count - 1
                If Val(Split(.Childs(i).Record(COL_主题序号).Record.Tag, "-")(1)) = 0 Then
                    .Childs(i).Record(COL_有效天数).Value = ""

                Else
                    .Childs(i).Record(COL_有效天数).Value = IIF(txtDays.Text = "", 0, txtDays.Text)
                End If
            Next i
        End With
        UnitReportControl.Populate
    End If
End Sub

Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)
    If Chr(KeyCode) = "'" Then KeyCode = 0
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub UnitReportControl_ColumnClick(ByVal Column As XtremeReportControl.IReportColumn)
    Call Arrange(Column.Index)
End Sub

Public Sub Arrange(Column As Long)
    UnitReportControl.SortOrder.DeleteAll
    UnitReportControl.SortOrder.Add UnitReportControl.Columns.Find(Column)
    UnitReportControl.SortOrder(0).SortAscending = Not UnitReportControl.SortOrder(0).SortAscending
    UnitReportControl.Populate
End Sub


Private Sub UnitReportControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
         If Not (UnitReportControl.FocusedRow Is Nothing) Then
            If Not UnitReportControl.FocusedRow.GroupRow And UnitReportControl.FocusedRow.Childs.Count = 0 Then
              Call UnitReportControl_RowDblClick(UnitReportControl.FocusedRow, UnitReportControl.FocusedRow.Record.Item(COL_主题序号))
            End If
        End If
    End If
End Sub

Private Sub UnitReportControl_MouseUp(Button As Integer, Shift As Integer, x As Long, y As Long)
'功能:弹出邮件菜单
    Dim cbrPopupBar As CommandBar
    Dim cbrPopupItem As CommandBarControl
    Dim cbrMenuBar As CommandBarControl
    Dim cbrControl As Object
    
    If Button <> 2 Then Exit Sub
    
    If cbsMain.ActiveMenuBar.Controls(2).Visible = False Then Exit Sub

    Set cbrMenuBar = cbsMain.ActiveMenuBar.Controls(2)
    Set cbrPopupBar = cbsMain.Add("弹出菜单", xtpBarPopup)
    For Each cbrControl In cbrMenuBar.CommandBar.Controls
        Set cbrPopupItem = cbrPopupBar.Controls.Add(xtpControlButton, cbrControl.ID, cbrControl.Caption)
        cbrPopupItem.BeginGroup = cbrControl.BeginGroup
    Next
    cbrPopupBar.ShowPopup
End Sub

Private Sub UnitReportControl_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If Not (Row Is Nothing) Then
        If Not Row.GroupRow And Row.Childs.Count = 0 And Val(Split(Row.Record(COL_主题序号).Record.Tag, "-")(1)) <> 0 Then
            Call cbsMain_Execute(cbsMain.FindControl(, conMenu_Edit_Modify, True, True))
        Else
            Call cbsMain_Execute(cbsMain.FindControl(, conMenu_Edit_ModifyParent, True, True))
        End If
    End If
End Sub


Private Sub UnitReportControl_SelectionChanged()
'-------------------------------------------------
'功能:根据ReportControl的选择列，提取对应的病区主题信息
'
'--------------------------------------------------
    Dim i As Integer
    
    txtInfo.Text = "": txtInfo.Tag = "": lblSet(7).Tag = "": lblSet(8).Tag = "": imaCustom.Text = "": imaCustom.Tag = ""
    lblSet(9).Tag = "": cbo标记.Tag = "": lblSet(1).Tag = "": txtName.Text = "": lblSet(4).Tag = "": txtDays.Text = "": chkSpecial.Value = 0
    
    On Error GoTo ErrHand
        With UnitReportControl.FocusedRow
            If Not UnitReportControl.FocusedRow Is Nothing Then
                If Not .GroupRow And .Childs.Count = 0 Then
                    If Val(Split(.Record(COL_主题序号).Record.Tag, "-")(1)) <> 0 Then
                        cbo标记.ListIndex = SetCboIndex(cbo标记, Val(Split(.Record(COL_主题序号).Record.Tag, "-")(0)))
                        lblSet(9).Tag = .Record(COL_原始主题).Value
                        lblSet(8).Tag = .Record(COL_说明).Value
                        txtInfo.Text = .Record(COL_说明).Value
                        lblSet(7).Tag = IIF(Val(.Record(COL_标注).Icon) <= 0, "0", Val(.Record(COL_标注).Icon))
                        If lblSet(7).Tag >= mlngImgIndex Then
                            imaCustom.ComboItems(Val(lblSet(7).Tag) - mlngImgIndex + 1).Selected = True
                        End If
                        Call SetControlEnable(fraInfo.Tag <> "")
                        Call SetFraResize
                    Else
                        UnitReportControl.FocusedRow = UnitReportControl.Rows(UnitReportControl.FocusedRow.Index - 1)
                    End If
                Else
                    lblSet(1).Tag = Split(.Childs(0).Record(COL_主题序号).GroupCaption, "-")(1)
                    txtName.Text = lblSet(1).Tag
                    lblSet(4).Tag = Val(.Childs(0).Record(COL_有效天数).Value)
                    txtDays.Text = lblSet(4).Tag
                    
                    txtName.Enabled = fraUnit.Tag <> ""
                    txtDays.Enabled = fraUnit.Tag <> ""
                    
                    txtName.BackColor = IIF(fraUnit.Tag <> "", UnEnable_Color, Enable_Color)
                    txtDays.BackColor = IIF(fraUnit.Tag <> "", UnEnable_Color, Enable_Color)
                    chkSpecial.Visible = (m病区ID = 0)
                    chkSpecial.Enabled = fraUnit.Tag <> "" And (m病区ID = 0)
                    If m病区ID = 0 Then
                        chkSpecial.Value = Val(.Childs(0).Record(COL_是否特殊).Value)
                    End If
                    chkSpecial.Tag = chkSpecial.Value
                    
                    Call SetFraResize(True)
                End If
            End If
        End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function SetCboIndex(ByVal objCbo As Object, ByVal intItemData As Integer) As Integer
'------------------------------------------------------------------------
'功能:根据itemdata的值获取cbo的Index
'------------------------------------------------------------------------
    Dim i As Integer
    Dim intIndex As Integer
    
    intIndex = -1
    
    For i = 0 To objCbo.ListCount - 1
        If Val(objCbo.ItemData(i)) = intItemData Then
           intIndex = i
           Exit For
        End If
    Next i
    
    SetCboIndex = intIndex
End Function

Private Function GetCboText(ByVal objCbo As Object, ByVal intItemData As Integer) As String
'------------------------------------------------------------------------
'功能:根据itemdata的值获取cbo的Index
'------------------------------------------------------------------------
    Dim i As Integer
    Dim strText As String
    
    strText = ""
    
    For i = 0 To objCbo.ListCount - 1
        If Val(objCbo.ItemData(i)) = intItemData Then
           strText = objCbo.Text
           Exit For
        End If
    Next i
    
    GetCboText = strText
End Function

Private Function CheckChange() As Boolean
'-----------------------------------------------------
'功能:修改时检查内容是否发生变化
'-----------------------------------------------------
    Dim blnChage As Boolean
    If fraInfo.Tag = "修改" Then
        If Val(lblSet(9).Tag) <> cbo标记.ListIndex Or lblSet(8).Tag <> txtInfo.Text Or _
            Val(lblSet(7).Tag) <> imaCustom.SelectedItem.Index - 1 + mlngImgIndex Then
            blnChage = True
        End If
    ElseIf fraUnit.Tag = "修改" Then
        If lblSet(1).Tag <> txtName.Text Or lblSet(4).Tag <> txtDays.Text Or Val(chkSpecial.Tag) <> Val(chkSpecial.Value) Then
            blnChage = True
        End If
    End If
    CheckChange = blnChage
End Function

Private Function CheckUseUnit(ByVal lngUnitID As Long, ByVal lngSubjectID As Long, ByVal lngTracerID As Long) As Boolean
'----------------------------------------------------------
'功能：检查改标记内容是否正在使用
'参数：lngUnitId 病区ID，lngSubjectID 主题序号 ，lngTracerID 标记序号
'----------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim blnTrue As Boolean
    Dim strSql
    On Error GoTo ErrHand
    
    If lngTracerID <> 0 Then
        strSql = "Select 1 From 病区标记记录" & _
            "   WHERE  " & IIF(lngUnitID = 0, " 主题病区Id IS NULL ", " 病区Id=[1] ") & " and 主题序号=[2] and 标记序号=[3] And RowNum<2"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "病区标记记录", lngUnitID, lngSubjectID, lngTracerID)
        If Not rsTmp.EOF Then
            blnTrue = True
            MsgBox "该标记内容目前改病区正在使用,请取消使用后在删除.", vbInformation, gstrSysName
        End If
    Else
        strSql = _
            " SELECT 1" & vbNewLine & _
            " FROM 病区标记内容 A,病区标记记录 B" & vbNewLine & _
            " WHERE  " & IIF(lngUnitID = 0, " B.主题病区Id IS NULL ", " A.病区ID=B.病区ID ") & " And A.主题序号=B.主题序号 And " & IIF(lngUnitID = 0, " A.病区ID IS NULL ", " A.病区ID=[1] ") & " And A.主题序号=[2]  " & vbNewLine & _
            " And RowNum<2"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "病区标记记录", lngUnitID, lngSubjectID)
        If Not rsTmp.EOF Then
            blnTrue = True
            MsgBox "该标记分类下的标记内容目前改病区正在使用,请取消使用后在删除.", vbInformation, gstrSysName
        End If
    End If
    CheckUseUnit = blnTrue
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetNewPreID(ByVal lng病区id As Long, ByVal lngPreVId As Long) As Long
'--------------------------------------------------------------------
'功能:提取某病区某主题下的标记序号
'参数:lng病区ID：病区ID ； lngPreVID ：主题序号
'--------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strSql As String
    Dim lngPreID As Long, i As Integer
    Dim arrPreID, blnFind As Boolean
    On Error GoTo ErrHand
    arrPreID = Array()
    strSql = _
        " select 标记序号" & _
        " From 病区标记内容" & _
        " Where " & IIF(lng病区id = 0, " 病区Id IS NULL ", " 病区Id=[1] ") & " and 主题序号=[2] order by 标记序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng病区id, lngPreVId)
    Do While Not rsTemp.EOF
        ReDim Preserve arrPreID(UBound(arrPreID) + 1)
        arrPreID(UBound(arrPreID)) = Val(rsTemp!标记序号 & "")
        rsTemp.MoveNext
    Loop
    For i = 0 To UBound(arrPreID)
        If Val(arrPreID(i)) > i + 1 Then
            lngPreID = i + 1
            blnFind = True
            Exit For
        End If
    Next
    If blnFind = False Then
        lngPreID = i + 1
    End If
    
    GetNewPreID = lngPreID
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetNewSubjectId(ByVal lng病区id As Long) As Long
'------------------------------------------------------------------------
'功能:新增标注分类时，提取某病区标记主题的新主题序号
'------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    Dim lngSubjectID As Long
    Dim arrSubJectID, i As Integer
    Dim blnFind As Boolean
    
    On Error GoTo ErrHand:
    strSql = _
        " select 主题序号,说明 from 病区标记内容" & _
        " where " & IIF(lng病区id = 0, " 病区Id IS NULL ", " 病区Id=[1] ") & " And 标记序号=0 Order by 主题序号"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "病区标记内容", lng病区id)
    
    arrSubJectID = Array()
    With rsTmp
        Do While Not .EOF
            ReDim Preserve arrSubJectID(UBound(arrSubJectID) + 1)
            arrSubJectID(UBound(arrSubJectID)) = Val("" & !主题序号)
            .MoveNext
        Loop
    End With
    
    For i = 0 To UBound(arrSubJectID)
        If Val(arrSubJectID(i)) > i + 1 Then
            lngSubjectID = i + 1
            blnFind = True
            Exit For
        End If
    Next
    If blnFind = False Then
        lngSubjectID = i + 1
    End If

    GetNewSubjectId = lngSubjectID
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Public Function IsEqualInfo(strName As String, Optional ByVal blnSubject As Boolean, Optional ByVal strKey As String = "") As Boolean
'同一病区下标记需要说明不能重复
    Dim StrInfo As String
    Dim blnAdd As Boolean
    On Error GoTo ErrHand
    If blnSubject = True Then
        mrsData.Filter = "标记序号=0"
    Else
        mrsData.Filter = "标记序号>0"
    End If
    Do While Not mrsData.EOF
        blnAdd = False
        If strKey = "" Then
            blnAdd = True
        Else
            If blnSubject = True Then
                If "" & mrsData!主题序号 <> strKey Then blnAdd = True
            Else
                If "" & mrsData!主题序号 & "-" & "" & mrsData!标记序号 <> strKey Then blnAdd = True
            End If
        End If
        If blnAdd = True Then StrInfo = StrInfo & "'" & mrsData!说明
        mrsData.MoveNext
    Loop
    If Left(StrInfo, 1) = "'" Then StrInfo = Mid(StrInfo, 2)
    '检查标记分类名称是否重复
    If InStr(1, "'" & StrInfo & "'", "'" & strName & "'") <> 0 Then
        If blnSubject = True Then
            MsgBox "此标记名称已经存在,请重新填写！", vbInformation, gstrSysName
        Else
            MsgBox "此标记说明已经存在,请重新填写！", vbInformation, gstrSysName
        End If
        Exit Function
    End If
    IsEqualInfo = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function CheckUnitSubject(ByVal lng病区id As Long) As Long
'---------------------------------------------------
'功能:检查是否存在标注主题名称,不存在提示操作员进行设置
'---------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    On Error GoTo ErrHand
    
    If lng病区id = 0 Then '公共图标
        strSql = " select 主题序号,说明 from 病区标记内容  where 病区Id is null and  标记序号=0"
    Else
        strSql = " select 主题序号,说明 from 病区标记内容  where 病区Id=[1] and  标记序号=0"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "病区标记内容", lng病区id)
    
    cbo标记.Clear
    With rsTmp
        Do While Not .EOF
            cbo标记.AddItem zlCommFun.Nvl(!说明, "个性标注" & zlCommFun.Nvl(!主题序号))
            cbo标记.ItemData(cbo标记.NewIndex) = Val(zlCommFun.Nvl(!主题序号))
            If cbo标记.ListIndex = -1 Then
                Call zlControl.CboSetIndex(cbo标记.hwnd, cbo标记.NewIndex)
            End If
        .MoveNext
        Loop
    End With
                
    CheckUnitSubject = rsTmp.RecordCount

    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

'################################################################################################################
'## 功能：  将数据从一个XtremeReportControl控件复制到VSFlexGrid，以便进行打印
'################################################################################################################
Private Function zlReportToVSFlexGrid(vfgList As VSFlexGrid, rptList As ReportControl) As Boolean
    '-------------------------------------------------
    '将全部组强制展开,复制数据表格
    Dim rptCol As ReportColumn
    Dim rptRcd As ReportRecord
    Dim rptItem As ReportRecordItem
    Dim rptRow As ReportRow
    Dim strGroupCaption As String
    
    Dim lngCol As Long, lngRow As Long
    
    On Error GoTo ErrHand:
    For Each rptRow In rptList.Rows
        If rptRow.GroupRow Then rptRow.Expanded = True
    Next
    
    With vfgList
        .Clear
        .Rows = rptList.Records.Count + 1
        .Cols = 0: .Cols = rptList.Columns.Count
        .FixedCols = rptList.GroupsOrder.Count
        
        '标题行复制
        .Row = 0
        lngCol = 0
        For Each rptCol In rptList.GroupsOrder
            .TextMatrix(0, lngCol) = rptCol.Caption
            .ColData(lngCol) = rptCol.ItemIndex
            Select Case rptCol.Alignment
            Case xtpAlignmentLeft: .FixedAlignment(lngCol) = flexAlignLeftCenter
            Case xtpAlignmentCenter: .FixedAlignment(lngCol) = flexAlignCenterCenter
            Case xtpAlignmentRight:  .FixedAlignment(lngCol) = flexAlignRightCenter
            End Select
            .Cell(flexcpAlignment, 0, lngCol, .FixedRows - 1) = flexAlignCenterCenter
            .Cell(flexcpAlignment, .FixedRows, lngCol, .Rows - 1) = .FixedAlignment(lngCol)
            .ColWidth(lngCol) = 100 * 15
            .MergeCol(lngCol) = True
            lngCol = lngCol + 1
        Next
        For Each rptCol In rptList.Columns
            If rptCol.Visible Then
                .TextMatrix(0, lngCol) = rptCol.Caption
                If rptCol.Caption = "标注" Then rptCol.Width = 10
                .ColData(lngCol) = rptCol.ItemIndex
                Select Case rptCol.Alignment
                Case xtpAlignmentLeft: .ColAlignment(lngCol) = flexAlignLeftCenter
                Case xtpAlignmentCenter: .ColAlignment(lngCol) = flexAlignCenterCenter
                Case xtpAlignmentRight: .ColAlignment(lngCol) = flexAlignRightCenter
                End Select
                .Cell(flexcpAlignment, 0, lngCol, .FixedRows - 1) = flexAlignCenterCenter
                .Cell(flexcpAlignment, .FixedRows, lngCol, .Rows - 1) = .ColAlignment(lngCol)
                If rptCol.Width < 20 Then
                    .ColWidth(lngCol) = 0
                Else
                    .ColWidth(lngCol) = rptCol.Width * 15
                End If
                lngCol = lngCol + 1
            End If
        Next
        vfgList.Cols = lngCol
        
        '数据行复制
        lngRow = 0
        For Each rptRow In rptList.Rows
            If rptRow.GroupRow = False Then
                lngRow = lngRow + 1
                For lngCol = 0 To .Cols - 1
                    If rptRow.Record(.ColData(lngCol)).GroupCaption <> "" Then
                        strGroupCaption = Split(rptRow.Record(.ColData(lngCol)).GroupCaption, "：")(1)
                    Else
                        strGroupCaption = rptRow.Record(.ColData(lngCol)).GroupCaption
                    End If
                    .TextMatrix(lngRow, lngCol) = IIF(.TextMatrix(0, lngCol) = "主题序号", strGroupCaption, rptRow.Record(.ColData(lngCol)).Value)
                    If rptRow.Record(.ColData(lngCol)).Icon > 0 Then
                        '.CellPicture = zlCommFun.GetPaitSignImageList(0).ListImages(rptRow.Record(.ColData(lngCol)).Icon).Picture
                    End If
                Next
            End If
        Next
    End With
    zlReportToVSFlexGrid = True
    Exit Function

ErrHand:
    zlReportToVSFlexGrid = False
End Function

Private Function ReDimArray(ByRef strArray() As String) As Long
    '----------------------------------------------------------------------
    '功能：重新定义数组
    '----------------------------------------------------------------------
    Dim lngCount As Long
    Dim strTmp As String
    
    On Error GoTo InitHand
    strTmp = strArray(0)
    lngCount = UBound(strArray) + 1
    GoTo OkHand
InitHand:
    lngCount = 1
OkHand:
    ReDim Preserve strArray(0 To lngCount)
    ReDimArray = lngCount
End Function

Private Sub Record_Init(ByRef rsObj As ADODB.Recordset, ByVal strFields As String)
    Dim arrFields, intField As Integer
    Dim strFieldName As String, intType As Integer, lngLength As Long
    '初始化映射记录集
    'strFields:字段名,类型,长度|字段名,类型,长度    如果长度为零,则取默认长度
    '字符型:adLongVarChar;数字型:adDouble;日期型:adDBDate
    
    '例子：
    'Dim rsVoucher As New ADODB.Recordset, strFields As String
    'strFields = "RecordID," & adDouble & ",18|科目ID," & adDouble & ",18|摘要, " & adLongVarChar & ",50|" & _
    '"删除," & adDouble & ",1"
    'Call Record_Init(rsVoucher, strFields)

    arrFields = Split(strFields, "|")
    Set rsObj = New ADODB.Recordset

    With rsObj
        If .State = 1 Then .Close
        For intField = 0 To UBound(arrFields)
            strFieldName = Split(arrFields(intField), ",")(0)
            intType = Split(arrFields(intField), ",")(1)
            lngLength = Split(arrFields(intField), ",")(2)

            '获取字段缺省长度
            If lngLength = 0 Then
                Select Case intType
                Case adDouble
                    lngLength = madDoubleDefault
                Case adVarChar
                    lngLength = madLongVarCharDefault
                Case adLongVarChar
                    lngLength = madLongVarCharDefault
                Case Else
                    lngLength = madDbDateDefault
                End Select
            End If
            .Fields.Append strFieldName, intType, lngLength, adFldIsNullable
        Next
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Private Sub Record_Add(ByRef rsObj As ADODB.Recordset, ByVal strFields As String, ByVal strValues As String)
    Dim arrFields, arrValues, intField As Integer
    '添加记录
    'strFields:字段名|字段名
    'strValues:值|值
    
    '例子：
    'Dim strFields As String, strValues As String
    'strFields = "RecordID|科目ID|摘要"
    'strValues = "5188|6666|科目名称"
    'Call Record_Update(rsVoucher, strFields, strValues)

    arrFields = Split(strFields, "|")
    arrValues = Split(strValues, "|")
    intField = UBound(arrFields)
    If intField = 0 Then Exit Sub

    With rsObj
        .AddNew
        For intField = 0 To intField
            .Fields(arrFields(intField)).Value = IIF(UCase(arrValues(intField)) = "NULL", Null, arrValues(intField))
        Next
        .Update
    End With
End Sub
