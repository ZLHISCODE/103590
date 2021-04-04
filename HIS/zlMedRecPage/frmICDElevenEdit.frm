VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmICDElevenEdit 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5355
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   10065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picICDEleven 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   5295
      Left            =   0
      ScaleHeight     =   5265
      ScaleWidth      =   10035
      TabIndex        =   0
      Top             =   0
      Width           =   10065
      Begin VB.CommandButton cmdAddExPand 
         Appearance      =   0  'Flat
         Caption         =   "增加扩展码"
         Height          =   375
         Left            =   1320
         TabIndex        =   8
         Top             =   4800
         Width           =   1100
      End
      Begin VB.CommandButton cmdAddMain 
         Appearance      =   0  'Flat
         Caption         =   "增加主干码"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   4800
         Width           =   1100
      End
      Begin VB.PictureBox picInfectInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   3495
         Left            =   0
         ScaleHeight     =   3465
         ScaleWidth      =   10035
         TabIndex        =   2
         Top             =   1200
         Visible         =   0   'False
         Width           =   10065
         Begin VB.ComboBox cboRelation 
            Height          =   300
            ItemData        =   "frmICDElevenEdit.frx":0000
            Left            =   1680
            List            =   "frmICDElevenEdit.frx":0002
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   120
            Width           =   8180
         End
         Begin VB.ListBox lstInfectParts 
            Appearance      =   0  'Flat
            Height          =   2340
            ItemData        =   "frmICDElevenEdit.frx":0004
            Left            =   240
            List            =   "frmICDElevenEdit.frx":000B
            Style           =   1  'Checkbox
            TabIndex        =   3
            Top             =   840
            Width           =   9615
         End
         Begin VB.Label lblBaseInfo 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "感染与死亡的关系"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   10
            Left            =   120
            TabIndex        =   6
            Top             =   180
            Width           =   1440
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "感染部位"
            Height          =   180
            Index           =   128
            Left            =   120
            TabIndex        =   5
            Top             =   480
            Width           =   720
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsDiagICDEleven 
         Height          =   1080
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   10065
         _cx             =   17754
         _cy             =   1905
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
         ForeColorSel    =   -2147483634
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
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmICDElevenEdit.frx":001F
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
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   2
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
      Begin VB.Image imgEixt 
         Height          =   360
         Left            =   9480
         Picture         =   "frmICDElevenEdit.frx":0114
         Top             =   4800
         Width           =   360
      End
      Begin VB.Image imgDelete 
         Height          =   360
         Left            =   8160
         Picture         =   "frmICDElevenEdit.frx":07FE
         Top             =   4800
         Width           =   360
      End
      Begin VB.Image imgDown 
         Height          =   360
         Left            =   7440
         Picture         =   "frmICDElevenEdit.frx":0EE8
         Top             =   4800
         Width           =   360
      End
      Begin VB.Image imgUp 
         Height          =   360
         Left            =   6840
         Picture         =   "frmICDElevenEdit.frx":15D2
         Top             =   4800
         Width           =   360
      End
      Begin VB.Image imgSave 
         Height          =   360
         Left            =   8880
         Picture         =   "frmICDElevenEdit.frx":1CBC
         Top             =   4800
         Width           =   360
      End
   End
End
Attribute VB_Name = "frmICDElevenEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mfrmParent As Object
Private mlngModel As Integer, mintDiagType As Integer
Private mRsDiag As ADODB.Recordset
Private mlngX As Long, mlngY As Long, mlngH As Long
Private mstrRelation As String
Private mlng科室ID As Long
Private mstr性别 As String
Private mblnEnter As Boolean
Private mint简码 As Integer
Private mstrLike As String
Private mlng病人ID As Long
Private mlng就诊ID As Long
Private mlngPatiType As Long
Private mDiagTag As String
Private mlng标记 As Long
Private mblnSave As Boolean
Private mblnShow As Boolean

Public Function ShowMe(ByRef frmParent As Object, ByVal lngModel As Long, ByVal lng病人ID As Long, ByVal lng就诊ID As Long, ByVal lngPatiType As Long, ByVal lng科室ID As Long, ByVal str性别 As String, ByVal intDiagType As Integer, ByRef rsDiag As ADODB.Recordset, ByVal X As Long, _
    ByVal Y As Long, ByVal txtH As Long, Optional ByRef strRelation As String) As Boolean
'功能：ICD-11诊断操作界面显示
'参数：lngModel-模块号 lng病人ID-病人ID lng就诊ID-(门诊病人“挂号ID”，住院病人“主页ID”) lngPatiType-病人类型（1-门诊病人 2-住院病人）
'      lng科室ID-出院科室ID str性别-性别 intDiagType-诊断类型（1-门诊诊断 2-入院诊断 3-出院诊断 5-院内感染 6-病理诊断 7-损伤中毒 11-中医门诊诊断 12-中医入院诊断 13-中医出院诊断）
'      rsDiag-intDiagType对应的诊断记录集 strRelation-感染与死亡关系及感染部位数据拼接的字符串（格式为xx|xx|xx&a）,当strRelation="-1"表示不显示感染与死亡关系及感染部位
    Set mfrmParent = frmParent
    mlngModel = lngModel
    mintDiagType = intDiagType
    Set mRsDiag = rsDiag
    mlngY = Y
    mlngX = X
    mlngH = txtH
    mstrRelation = strRelation
    mlng科室ID = lng科室ID
    mstr性别 = str性别
    mlng病人ID = lng病人ID
    mlng就诊ID = lng就诊ID
    mlngPatiType = lngPatiType
    Me.Show 1, mfrmParent
    strRelation = mstrRelation
    Set rsDiag = mRsDiag
    If mblnSave Then
        ShowMe = True
    Else
        ShowMe = False
    End If
    mblnSave = False
    mblnShow = False
End Function

Private Sub cmdAddExPand_Click()
'功能：增加扩展码
    Dim LngRow As Long
    Dim i As Long, j As Long
    Dim blnAdd As Boolean
    Dim k As Long
    
    With vsDiagICDEleven
        blnAdd = True
        vsDiagICDEleven.SetFocus
        If .TextMatrix(.Row, Eleven_诊断类型) = "主干码" Then
            k = 0
            j = 0
        Else
            k = 1
            j = 1
            If .TextMatrix(.Row, Eleven_诊断描述) = "" Then blnAdd = False: Exit Sub
        End If
        '如果主干码下对应的扩展码行有空行的情况下就不能再增加主干码
        For i = .Row To .Rows - 1
            If i + 1 <= .Rows - 1 Then
                If .TextMatrix(i + 1, Eleven_诊断类型) = "扩展码" Or .TextMatrix(i + 1, Eleven_诊断类型) = "证  候" Then
                    j = j + 1
                Else
                    Exit For
                End If
                If .TextMatrix(i + 1, Eleven_诊断描述) = "" Then blnAdd = False: Exit For
            End If
        Next
        If Not blnAdd Then Exit Sub
        If k <> 0 Then
            For i = .Row To .FixedRows Step -1
                If i - 1 >= .FixedRows Then
                    If .TextMatrix(i - 1, Eleven_诊断类型) = "扩展码" Or .TextMatrix(i - 1, Eleven_诊断类型) = "证  候" Then
                        j = j + 1
                    Else
                        Exit For
                    End If
                    If .TextMatrix(i - 1, Eleven_诊断描述) = "" Then blnAdd = False: Exit For
                End If
            Next
        End If
        If Not blnAdd Then Exit Sub
        '主干码对应的扩展码不能超过9条数据
        If j < 99 Then
            LngRow = .Row + 1: .AddItem "", LngRow
            .TextMatrix(LngRow, Eleven_诊断类型) = IIf(mDiagTag = "西医", "扩展码", "证  候")
            .Cell(flexcpBackColor, LngRow, .FixedRows, LngRow, .Cols - 1) = GRD_UNEDITCELL_COLOR
            .Row = LngRow: .Col = Eleven_诊断描述
            vsDiagICDEleven.SetFocus
        Else
            MsgBox "主干码对应的扩展码诊断不能超过9条！", vbInformation, gstrSysName
            Exit Sub
        End If
    End With
    Call ChangeVSHeight
End Sub

Private Sub cmdAddMain_Click()
'功能增加主干码
    Dim i As Long, j As Long
    Dim LngRow
    Dim blnAdd As Boolean
    
    With vsDiagICDEleven
        blnAdd = True
        vsDiagICDEleven.SetFocus
        '如果诊断列表中存在诊断描述为空的行则不能增加主干码
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, Eleven_诊断类型) = "主干码" Then
                j = j + 1
                If .TextMatrix(i, Eleven_诊断描述) = "" Then blnAdd = False: Exit For
            End If
        Next
        If Not blnAdd Then Exit Sub
        '同类型诊断的主干码不能超过9诊断
        If j < 99 Then
            LngRow = .Rows: .AddItem "", LngRow
            .TextMatrix(LngRow, Eleven_诊断类型) = "主干码"
            .Cell(flexcpBackColor, LngRow, .FixedRows, LngRow, .Cols - 1) = &HC0FFC0
            Call ChangeVSHeight
            LngRow = .Rows: .AddItem "", LngRow
            .TextMatrix(LngRow, Eleven_诊断类型) = IIf(mDiagTag = "西医", "扩展码", "证  候")
            .Cell(flexcpBackColor, LngRow, .FixedRows, LngRow, .Cols - 1) = GRD_UNEDITCELL_COLOR
            Call ChangeVSHeight
            vsDiagICDEleven.SetFocus
            .Row = LngRow - 1: .Col = Eleven_诊断描述
            .ShowCell .Row, .Col
        Else
            MsgBox "主干码诊断不能超过9条！", vbInformation, gstrSysName
            Exit Sub
        End If
    End With
    Call ChangeVSHeight
End Sub

Private Sub Form_Activate()
    Call ChangeVSHeight
    vsDiagICDEleven.Row = vsDiagICDEleven.FixedRows
    vsDiagICDEleven.Col = Eleven_诊断描述
    vsDiagICDEleven.ShowCell vsDiagICDEleven.Row, vsDiagICDEleven.Col
    vsDiagICDEleven.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Dim lngScrH As Long
    Me.KeyPreview = True
    If mRsDiag Is Nothing Then Call InitRsICDEleven(mRsDiag)
    If mRsDiag.Fields.Count = 0 Then Call InitRsICDEleven(mRsDiag)
    '只有住院首页和病案首页才会显示感染与死亡关系和感染部位
    If (mlngModel = p住院医生站 Or mlngModel = p病案管理) And mintDiagType = 5 And mRsDiag.RecordCount >= 1 And mstrRelation <> "-1" Then
        mblnShow = True
    Else
        mblnShow = False
    End If
    If mstrRelation = "-1" Then
        mstrRelation = ""
    End If
    picICDEleven.Height = IIf(mblnShow = False, picICDEleven.Height - picInfectInfo.Height - 100, picICDEleven.Height)
    Me.Height = picICDEleven.Height
    Me.Left = mlngX
    lngScrH = GetSystemMetrics(SM_CYFULLSCREEN) * 15 '屏幕可用高度
    If mlngY + mlngH + Me.Height > lngScrH Then
        Me.Top = mlngY - Me.Height
    Else
        Me.Top = mlngY + mlngH
    End If
    '初始化表格
    Call InitTable(vsDiagICDEleven)
    '初始化界面数据
    Call InitData
    '加载诊断
    Call LoadData
End Sub

Private Sub InitTable(ByRef vsTmp As VSFlexGrid)
'功能：初始化表格
    Dim strHeader As String, strRows As String
    Dim LngRow As Long
    
    With vsTmp
        strHeader = "诊断类别,1095,4;诊断编码,1905,4;诊断描述,6610,1;=疾病ID;医嘱IDs;是否病人"
        strRows = Eleven_诊断类型 & ",主干码;" & Eleven_诊断类型 & ",扩展码"
        Call Grid.Init(vsTmp, strHeader, strRows, 1, 1)
        .Rows = 3
        .Cell(flexcpBackColor, .FixedRows, Eleven_诊断编码, .Rows - 1, Eleven_诊断描述) = GRD_UNEDITCELL_COLOR     '灰蓝色
        LngRow = .FindRow("主干码", , Eleven_诊断类型, True)
        .Cell(flexcpBackColor, LngRow, .FixedRows, LngRow, .Cols - 1) = &HC0FFC0
        LngRow = .FindRow("扩展码", , Eleven_诊断类型, True)
        .Cell(flexcpBackColor, LngRow, .FixedRows, LngRow, .Cols - 1) = GRD_UNEDITCELL_COLOR
        .Row = .FixedRows: .Col = Eleven_诊断描述
    End With
End Sub

Private Sub Form_Resize()
    picInfectInfo.Visible = mblnShow
End Sub

Private Sub imgDelete_Click()
'功能：删除当前行
    Dim LngRow As Long
    Dim strMsg As String
    Dim lng医嘱ID As Long
    Dim rsTmp As ADODB.Recordset
    Dim str类型 As String
    Dim i As Long, j As Long
    
    With vsDiagICDEleven
        LngRow = .Row
        If LngRow + 1 <= .Rows - 1 And .Rows > 3 Then
            If .TextMatrix(LngRow, Eleven_诊断类型) = "主干码" Then
            '如果当前行是主干码则删除主干码及对应的扩展码
                For i = LngRow To .Rows - 1
                    If i <= .Rows - 1 Then
                        If .Rows > 3 Then
                            str类型 = .TextMatrix(i, Eleven_诊断类型)
                            .Cell(flexcpText, i, .FixedCols, i, .Cols - 1) = ""
                            .Cell(flexcpData, i, .FixedCols, i, .Cols - 1) = Empty
                            .TextMatrix(i, Eleven_诊断类型) = str类型
                            If i + 1 <= .Rows - 1 Then
                                For j = Eleven_诊断类型 To .Cols - 1
                                    .TextMatrix(i, j) = .TextMatrix(i + 1, j)
                                    .Cell(flexcpData, i, j) = .Cell(flexcpData, i + 1, j)
                                Next
                                .Cell(flexcpBackColor, i, .FixedRows, i, .Cols - 1) = .Cell(flexcpBackColor, i + 1, .FixedRows, i + 1, .Cols - 1)
                                .RowData(i) = .RowData(i + 1)
                                If .TextMatrix(i + 1, Eleven_诊断类型) = "主干码" Then
                                    .RemoveItem i + 1
                                    Exit For
                                Else
                                    .RemoveItem i + 1
                                    i = i - 1
                                End If
                            Else
                                .RemoveItem i
                            End If
                        Else
                            str类型 = .TextMatrix(i, Eleven_诊断类型)
                            .Cell(flexcpText, i, .FixedCols, i, .Cols - 1) = ""
                            .Cell(flexcpData, i, .FixedCols, i, .Cols - 1) = Empty
                            .TextMatrix(i, Eleven_诊断类型) = str类型
                            .Cell(flexcpBackColor, i, .FixedCols, i, .Cols - 1) = IIf(i = 1, &HC0FFC0, GRD_UNEDITCELL_COLOR)
                        End If
                    End If
                Next
            Else
                '如果当前行是扩展码则删除当前行
                str类型 = .TextMatrix(LngRow, Eleven_诊断类型)
                .Cell(flexcpText, LngRow, .FixedCols, LngRow, .Cols - 1) = ""
                .Cell(flexcpData, LngRow, .FixedCols, LngRow, .Cols - 1) = Empty
                .TextMatrix(LngRow, Eleven_诊断类型) = str类型
                If .TextMatrix(LngRow - 1, Eleven_诊断类型) = "主干码" Then
                    If LngRow + 1 <= .Rows - 1 Then
                        If .TextMatrix(LngRow + 1, Eleven_诊断类型) = "扩展码" Or .TextMatrix(LngRow + 1, Eleven_诊断类型) = "证  候" Then
                            .RemoveItem LngRow
                        End If
                    End If
                Else
                    For j = Eleven_诊断类型 To .Cols - 1
                        .TextMatrix(LngRow, j) = .TextMatrix(LngRow + 1, j)
                        .Cell(flexcpData, i, j) = .Cell(flexcpData, LngRow + 1, j)
                    Next
                    .Cell(flexcpBackColor, LngRow, .FixedRows, LngRow, .Cols - 1) = .Cell(flexcpBackColor, LngRow + 1, .FixedRows, LngRow + 1, .Cols - 1)
                    .RowData(LngRow) = .RowData(LngRow + 1)
                    .RemoveItem LngRow + 1
                End If
            End If
        Else
            If .Rows > 3 Then
                For i = LngRow To .Rows - 1
                    If i <= .Rows - 1 Then
                        If i - 1 <= .Rows - 1 Then
                            If .TextMatrix(i - 1, Eleven_诊断类型) = "主干码" Then
                                 str类型 = .TextMatrix(i, Eleven_诊断类型)
                                .Cell(flexcpText, i, .FixedCols, i, .Cols - 1) = ""
                                .Cell(flexcpData, i, .FixedCols, i, .Cols - 1) = Empty
                                .TextMatrix(i, Eleven_诊断类型) = str类型
                                 .Cell(flexcpBackColor, i, .FixedCols, i, .Cols - 1) = IIf(i = 1, &HC0FFC0, GRD_UNEDITCELL_COLOR)
                            Else
                                .RemoveItem i
                                i = i - 1
                            End If
                        Else
                            .RemoveItem i
                            i = i - 1
                        End If
                    End If
                Next
            Else
                For i = LngRow To .Rows - 1
                    str类型 = .TextMatrix(i, Eleven_诊断类型)
                    .Cell(flexcpText, i, .FixedCols, i, .Cols - 1) = ""
                    .Cell(flexcpData, i, .FixedCols, i, .Cols - 1) = Empty
                    .TextMatrix(i, Eleven_诊断类型) = str类型
                    .Cell(flexcpBackColor, i, .FixedCols, i, .Cols - 1) = IIf(i = 1, &HC0FFC0, GRD_UNEDITCELL_COLOR)
                Next
            End If
        End If
    End With
    Call ChangeVSHeight
End Sub

Private Sub imgDown_Click()
'功能：当前行向下移
    With vsDiagICDEleven
        '如果当前行是扩展码且下一行为主干码则不能再往下移动
        If .Row + 1 <= .Rows - 1 And (.TextMatrix(.Row, Eleven_诊断类型) = "扩展码" Or .TextMatrix(.Row, Eleven_诊断类型) = "证  候") Then
            If .TextMatrix(.Row + 1, Eleven_诊断类型) = "主干码" Then Exit Sub
        End If
        '向下移动
        Call MoveCurrRow(.TextMatrix(.Row, Eleven_诊断类型), .Row, -1)
    End With
End Sub

Private Sub imgEixt_Click()
    Unload Me
End Sub

Private Sub imgSave_Click()
'保存界面录入的数据
    If Not CheckDate Then Exit Sub
    If Not SaveData Then
        Exit Sub
    Else
        Unload Me
    End If
End Sub

Private Function CheckDate() As Boolean
'检查诊断列表录入的数据
    Dim i As Long, j As Long
    Dim blnHaveDaig As Boolean
    Dim str诊断类型 As String
    Dim lngColor As Long
    
    With vsDiagICDEleven
        '检查是否存在两行相同的主干码诊断
        For i = .FixedRows To .Rows - 1
            If Trim(.TextMatrix(i, Eleven_诊断描述)) <> "" And .TextMatrix(i, Eleven_诊断类型) = "主干码" Then
                If i <> .Rows - 1 Then
                     For j = i + 1 To .Rows - 1
                        If .TextMatrix(j, Eleven_诊断类型) = "主干码" Then
                            If Trim(.TextMatrix(i, Eleven_诊断描述)) = Trim(.TextMatrix(j, Eleven_诊断描述)) Then
                                .Row = i: .Col = Eleven_诊断描述
                                lngColor = .CellBackColor: .CellBackColor = &HC0C0FF
                                Call .ShowCell(.Row, .Col)
                                MsgBox "发现存在两行相同的诊断信息。", vbInformation, gstrSysName
                                .CellBackColor = lngColor
                                str诊断类型 = .TextMatrix(i, Eleven_诊断类型)
                                vsDiagICDEleven.SetFocus: Exit Function
                            End If
                        End If
                     Next
                End If
            End If
        Next
        '检查主干码对应的扩展码是否存在两行相同的诊断
        For i = .FixedRows To .Rows - 1
            If Trim(.TextMatrix(i, Eleven_诊断描述)) <> "" And (.TextMatrix(i, Eleven_诊断类型) = "扩展码" Or .TextMatrix(i, Eleven_诊断类型) = "证  候") Then
                If i <> .Rows - 1 Then
                     For j = i + 1 To .Rows - 1
                        If .TextMatrix(j, Eleven_诊断类型) = "主干码" Then
                            Exit For
                        Else
                            If Trim(.TextMatrix(i, Eleven_诊断描述)) = Trim(.TextMatrix(j, Eleven_诊断描述)) Then
                                .Row = i: .Col = Eleven_诊断描述
                                lngColor = .CellBackColor: .CellBackColor = &HC0C0FF
                                Call .ShowCell(.Row, .Col)
                                MsgBox "发现存在两行相同的诊断信息", vbInformation, gstrSysName
                                .CellBackColor = lngColor
                                str诊断类型 = .TextMatrix(i, Eleven_诊断类型)
                                vsDiagICDEleven.SetFocus: Exit Function
                            End If
                        End If
                     Next
                End If
            End If
        Next
        '检查扩展码是否对应了主干码
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, Eleven_诊断类型) = "主干码" Then
                If i <> .Rows - 1 Then
                     For j = i + 1 To .Rows - 1
                        If .TextMatrix(j, Eleven_诊断类型) = "主干码" Then
                            Exit For
                        Else
                            If .TextMatrix(j, Eleven_诊断描述) <> "" Then blnHaveDaig = True
                        End If
                    Next
                    If blnHaveDaig And .TextMatrix(i, Eleven_诊断描述) = "" Then
                        .Row = i: .Col = Eleven_诊断描述
                        lngColor = .CellBackColor: .CellBackColor = &HC0C0FF
                        Call .ShowCell(.Row, .Col)
                        MsgBox "扩展码没有对应的主干码，请检查！", vbInformation, gstrSysName
                        .CellBackColor = lngColor
                        str诊断类型 = .TextMatrix(i, Eleven_诊断类型)
                        vsDiagICDEleven.SetFocus: Exit Function
                    End If
                End If
            End If
        Next
    End With
    CheckDate = True
End Function

Private Function SaveData() As Boolean
    '保存界面录入的数据
    Dim strValues As String
    Dim j As Long, i As Long
    Dim k As Long
    Dim rsDiag As ADODB.Recordset
    
    Call InitRsICDEleven(rsDiag)
    
    With vsDiagICDEleven
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, Eleven_诊断描述) <> "" Then
                If .TextMatrix(i, Eleven_诊断类型) = "主干码" Then
                    k = 0
                    j = j + 1
                End If
                If .TextMatrix(i, Eleven_诊断类型) = "扩展码" Or .TextMatrix(i, Eleven_诊断类型) = "证  候" Then
                    k = k + 1
                End If
                rsDiag.AddNew Array("信息名", "诊断类型", "诊断编码", "诊断描述", "疾病ID", "医嘱IDs", "Tag", "IndexEx", "标记"), Array("ICD-11", .TextMatrix(i, Eleven_诊断类型), .TextMatrix(i, Eleven_诊断编码), .TextMatrix(i, Eleven_诊断描述), Val(.TextMatrix(i, Eleven_疾病ID)), .TextMatrix(i, Eleven_医嘱IDs), mDiagTag, IIf(.TextMatrix(i, Eleven_诊断类型) = "主干码", IIf(j <= 9, "0" & j, "" & j), IIf(j <= 9, "0" & j, "" & j) & IIf(k <= 9, "0" & k, "" & k)), mlng标记)
            End If
        Next
        Set mRsDiag = zlDatabase.CopyNewRec(rsDiag)
    End With
    '感染与死亡关系及感染部位
    If picInfectInfo.Visible Then
        For j = 0 To lstInfectParts.ListCount - 1
            If lstInfectParts.Selected(j) = True Then
                strValues = strValues & "|" & lstInfectParts.ItemData(j)
            End If
        Next
        If strValues <> "" Then
            strValues = Mid(strValues, 2)
        End If
        strValues = strValues & "&" & zlcommfun.GetNeedName(cboRelation.Text, "-")
    End If
    mstrRelation = strValues
    SaveData = True
    mblnSave = SaveData
End Function

Private Sub imgUp_Click()
'向上移动
    With vsDiagICDEleven
        '如果当前行是扩展码且上一行为主干码则不能向上移动
        If .Row - 1 >= .FixedRows And (.TextMatrix(.Row, Eleven_诊断类型) = "扩展码" Or .TextMatrix(.Row, Eleven_诊断类型) = "证  候") Then
            If .TextMatrix(.Row - 1, Eleven_诊断类型) = "主干码" Then Exit Sub
        End If
        '向上移动
        Call MoveCurrRow(.TextMatrix(.Row, Eleven_诊断类型), .Row, 1)
    End With
End Sub

Private Sub picICDEleven_Resize()
    If picInfectInfo.Visible Then
        picInfectInfo.Top = vsDiagICDEleven.Top + vsDiagICDEleven.Height + 100
        cmdAddMain.Top = picInfectInfo.Top + picInfectInfo.Height + 100
        cmdAddExPand.Top = picInfectInfo.Top + picInfectInfo.Height + 100
        imgUp.Top = picInfectInfo.Top + picInfectInfo.Height + 100
        imgDown.Top = picInfectInfo.Top + picInfectInfo.Height + 100
        imgDelete.Top = picInfectInfo.Top + picInfectInfo.Height + 100
        imgSave.Top = picInfectInfo.Top + picInfectInfo.Height + 100
        imgEixt.Top = picInfectInfo.Top + picInfectInfo.Height + 100
    Else
        cmdAddMain.Top = vsDiagICDEleven.Top + vsDiagICDEleven.Height + 100
        cmdAddExPand.Top = vsDiagICDEleven.Top + vsDiagICDEleven.Height + 100
        imgUp.Top = vsDiagICDEleven.Top + vsDiagICDEleven.Height + 100
        imgDown.Top = vsDiagICDEleven.Top + vsDiagICDEleven.Height + 100
        imgDelete.Top = vsDiagICDEleven.Top + vsDiagICDEleven.Height + 100
        imgSave.Top = vsDiagICDEleven.Top + vsDiagICDEleven.Height + 100
        imgEixt.Top = vsDiagICDEleven.Top + vsDiagICDEleven.Height + 100
    End If
End Sub

Private Sub LoadData()
'功能：显示数据
    Dim i As Long, j As Long, n As Long
    Dim lngRows As Long
    Dim strArrRelation As Variant
    Dim strTmp As String
    
    With vsDiagICDEleven
        For i = .FixedRows To .Rows - 1
            If i <= .Rows - 1 Then
                If i <> 1 And i <> 2 Then
                 .RemoveItem i
                End If
            End If
        Next
        For i = .FixedRows To .Rows - 1
            If i = 1 Then
                .TextMatrix(i, Eleven_诊断类型) = "主干码"
                .Cell(flexcpBackColor, i, .FixedRows, i, .Cols - 1) = &HC0FFC0
            Else
                .TextMatrix(i, Eleven_诊断类型) = "扩展码"
                .Cell(flexcpBackColor, i, .FixedRows, i, .Cols - 1) = GRD_UNEDITCELL_COLOR
            End If
            .TextMatrix(i, Eleven_疾病ID) = ""
            .TextMatrix(i, Eleven_诊断描述) = ""
            .Cell(flexcpData, i, Eleven_诊断描述) = .TextMatrix(i, Eleven_诊断描述)
            .TextMatrix(i, Eleven_诊断编码) = ""
            .TextMatrix(i, Eleven_医嘱IDs) = ""
            .TextMatrix(i, Eleven_是否病人) = ""
            .RowData(i) = ""
        Next
        If Not mRsDiag.EOF Then
            For i = 0 To mRsDiag.RecordCount - 1
                mDiagTag = mRsDiag!Tag
                If .TextMatrix(n, Eleven_诊断类型) = "主干码" And "" & mRsDiag!诊断类型 = "主干码" Then
                    n = n + 1
                    .AddItem "", n
                    .TextMatrix(n, Eleven_诊断类型) = IIf(mDiagTag = "西医", "扩展码", "证  候")
                    .Cell(flexcpBackColor, n, .FixedRows, n, .Cols - 1) = GRD_UNEDITCELL_COLOR
                End If
                n = n + 1
                If n > .Rows - 1 Then
                    .AddItem "", n
                End If
                .TextMatrix(n, Eleven_诊断类型) = "" & mRsDiag!诊断类型
                .TextMatrix(n, Eleven_诊断编码) = "" & mRsDiag!诊断编码
                .TextMatrix(n, Eleven_诊断描述) = "" & mRsDiag!诊断描述
                .TextMatrix(n, Eleven_疾病ID) = "" & mRsDiag!疾病id
                .TextMatrix(n, Eleven_医嘱IDs) = "" & mRsDiag!医嘱IDs
                .TextMatrix(n, Eleven_是否病人) = ""
                If "" & mRsDiag!诊断类型 = "主干码" Then
                    .Cell(flexcpBackColor, n, .FixedRows, n, .Cols - 1) = &HC0FFC0
                Else
                    .Cell(flexcpBackColor, n, .FixedRows, n, .Cols - 1) = GRD_UNEDITCELL_COLOR
                End If
                mlng标记 = Val("" & mRsDiag!标记)
                mRsDiag.MoveNext
                Call ChangeVSHeight
            Next
        End If
        If mintDiagType = 11 Or mintDiagType = 12 Or mintDiagType = 13 Then
            mDiagTag = "中医"
        Else
            mDiagTag = "西医"
        End If
        For i = .FixedRows To .Rows - 1
            If i + 1 <= .Rows - 1 Then
                If .TextMatrix(i, Eleven_诊断类型) = "主干码" Then
                    If .TextMatrix(i + 1, Eleven_诊断类型) <> "扩展码" And .TextMatrix(i + 1, Eleven_诊断类型) <> "证  候" Then
                        .AddItem "", i + 1
                        .TextMatrix(i + 1, Eleven_诊断类型) = IIf(mDiagTag = "西医", "扩展码", "证  候")
                        .Cell(flexcpBackColor, i + 1, .FixedRows, i + 1, .Cols - 1) = GRD_UNEDITCELL_COLOR
                    End If
                End If
            Else
                If .TextMatrix(i, Eleven_诊断类型) = "主干码" Then
                    .AddItem "", i + 1
                    .TextMatrix(i + 1, Eleven_诊断类型) = IIf(mDiagTag = "西医", "扩展码", "证  候")
                    .Cell(flexcpBackColor, i + 1, .FixedRows, i + 1, .Cols - 1) = GRD_UNEDITCELL_COLOR
                End If
            End If
            If .TextMatrix(i, Eleven_诊断类型) = "扩展码" Then
                .TextMatrix(i, Eleven_诊断类型) = IIf(mDiagTag = "西医", "扩展码", "证  候")
            End If
        Next
    End With
    
    If mstrRelation <> "" Then
        strTmp = Mid(mstrRelation, 1, InStr(mstrRelation, "&") - 1)
        If strTmp <> "" Then
            With lstInfectParts
                strArrRelation = Split(strTmp, "|")
                For j = 0 To .ListCount - 1
                    For i = LBound(strArrRelation) To UBound(strArrRelation)
                        If .ItemData(j) = strArrRelation(i) Then
                            .Selected(j) = True: Exit For
                        End If
                    Next
                Next
                .ListIndex = -1
            End With
        End If
        strTmp = Mid(mstrRelation, InStr(mstrRelation, "&") + 1)
        If strTmp <> "" Then
            Call Cbo.SeekIndex(cboRelation, strTmp)
        End If
    End If
End Sub

Private Sub InitData()
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    
    mint简码 = Val(zlDatabase.GetPara("简码方式"))
    mstrLike = IIf(zlDatabase.GetPara("输入匹配") = "0", "%", "")
    
    cboRelation.AddItem " "
    cboRelation.AddItem "0-直接"
    cboRelation.AddItem "1-间接"
    cboRelation.AddItem "2-无"
    cboRelation.ListIndex = -1
    
    strSql = "Select RowNum As ID, 编码, 简码, 名称, 缺省标志 缺省 From 感染部位"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "读取感染部位")
    With lstInfectParts
        If Not rsTmp.EOF Then
            lstInfectParts.Clear
            rsTmp.Sort = "编码,名称"
            Do While Not rsTmp.EOF
                .AddItem rsTmp!名称
                .ItemData(.NewIndex) = Val(rsTmp!编码)
                rsTmp.MoveNext
            Loop
        End If
    End With
End Sub

Private Function ChangeVSHeight() As Boolean
'功能：在PictureBox 上面调整VSF的大小
    Dim i As Long
    Dim lngOldVSFHeight As Long
    Dim lngRows As Long
    Dim lngVSFHeight As Long
    Dim lngRowHeight As Long
    Dim lngMaxHeight As Long
    Dim lngShowRows As Long
    Dim lngScrH As Long
    Dim lngLastHeight As Long
    
    lngRowHeight = IIf(vsDiagICDEleven.RowHeightMax < vsDiagICDEleven.RowHeightMin, vsDiagICDEleven.RowHeightMin, vsDiagICDEleven.RowHeightMax)
    lngOldVSFHeight = vsDiagICDEleven.Height
    lngRows = vsDiagICDEleven.Rows
    For i = 0 To vsDiagICDEleven.Rows - 1
        lngVSFHeight = lngVSFHeight + vsDiagICDEleven.RowHeight(i)
        lngShowRows = lngShowRows + 1
    Next
    lngVSFHeight = IIf(lngVSFHeight < lngShowRows * lngRowHeight, lngShowRows * lngRowHeight + 30, lngVSFHeight)
    lngScrH = GetSystemMetrics(SM_CYFULLSCREEN) * 15 '屏幕可用高度
    lngLastHeight = picICDEleven.Height + (lngVSFHeight - lngOldVSFHeight)
    If lngLastHeight > mlngY - 50 Then
        If mlngH + lngLastHeight > lngScrH - mlngY - mlngH Then
            vsDiagICDEleven.Height = vsDiagICDEleven.Height
        Else
            vsDiagICDEleven.Height = lngVSFHeight
        End If
    Else
        vsDiagICDEleven.Height = lngVSFHeight
    End If
    If vsDiagICDEleven.Height - lngOldVSFHeight <> 0 Then
        picICDEleven.Height = picICDEleven.Height + (vsDiagICDEleven.Height - lngOldVSFHeight)
    End If
    Me.Height = picICDEleven.Height
    If mlngY + mlngH + Me.Height > lngScrH Then
        Me.Top = mlngY - Me.Height
    Else
        Me.Top = mlngY + mlngH
    End If
End Function

Private Sub vsDiagICDEleven_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long, j As Long, k As Long
    Dim LngRow As Long, LngCol As Long
    Dim blnDel As Boolean
    
    LngRow = Row
    LngCol = Col
    With vsDiagICDEleven
        If LngCol = Eleven_诊断描述 Then
            ' .EditText = "" 排除单元格有内容并按回车的状况
            If (LngCol = Eleven_诊断描述 And .TextMatrix(LngRow, Eleven_诊断编码) <> "" Or LngCol = Eleven_诊断编码 And .TextMatrix(LngRow, Eleven_诊断描述) <> "") And .EditText = "" And .Cell(flexcpData, LngRow, LngCol) <> "" Then
                If .TextMatrix(LngRow, Eleven_诊断类型) = "主干码" Then
                    For i = LngRow + 1 To .Rows - 1
                        If .TextMatrix(i, Eleven_诊断类型) = "主干码" Then
                            Exit For
                        Else
                            If .TextMatrix(i, Eleven_诊断描述) <> "" Then
                                If MsgBox("是否在删除主干码的同时删除对应的扩展码？点击是，同步删除对应的扩展码；点击否，则不删对应除扩展码。", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then blnDel = True
                                Exit For
                            End If
                        End If
                    Next
                End If
                If Not blnDel Then
                    .TextMatrix(LngRow, LngCol) = .Cell(flexcpData, LngRow, LngCol)
                    .Cell(flexcpText, LngRow, .FixedCols, LngRow, .Cols - 1) = ""
                    .Cell(flexcpData, LngRow, .FixedCols, LngRow, .Cols - 1) = Empty
                    For j = .FixedCols To .Cols - 1
                        .TextMatrix(LngRow, j) = ""
                    Next
                Else
                    For i = LngRow To .Rows - 1
                        If i <= .Rows - 1 Then
                            k = k + 1
                            If .TextMatrix(i, Eleven_诊断类型) = "主干码" And i <> LngRow Then Exit For
                            .TextMatrix(i, LngCol) = .Cell(flexcpData, i, LngCol)
                            .Cell(flexcpText, i, .FixedCols, i, .Cols - 1) = ""
                            .Cell(flexcpData, i, .FixedCols, i, .Cols - 1) = Empty
                            For j = .FixedCols To .Cols - 1
                                .TextMatrix(i, j) = ""
                            Next
                            If k > 2 Then
                                .RemoveItem i
                                i = i - 1
                            End If
                        End If
                    Next
                End If
            End If
        End If
        Call ChangeVSHeight
        Call vsDiagICDEleven_AfterRowColChange(-1, -1, .Row, .Col)
    End With
End Sub

Private Sub vsDiagICDEleven_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim lngNewRow As Long, lngNewCol As Long
    
    lngNewRow = NewRow: lngNewCol = NewCol
    If lngNewRow = -1 Or lngNewCol = -1 Then Exit Sub
    If vsDiagICDEleven.Editable = flexEDNone Then Exit Sub
    
    With vsDiagICDEleven
        If Not ICDElevenEditable(lngNewRow, lngNewCol) Then
            .ComboList = ""
            .FocusRect = flexFocusLight
        Else
            .ComboList = ""
            .FocusRect = flexFocusSolid
            Select Case lngNewCol
                Case Eleven_诊断描述
                    .ComboList = "..."
                Case Eleven_诊断编码
                    
                Case Else
                    .ComboList = ""
            End Select
        End If
    End With
End Sub

Private Sub vsDiagICDEleven_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not ICDElevenEditable(Row, Col) Then
        Cancel = True
    End If
End Sub

Private Sub vsDiagICDEleven_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> Eleven_诊断描述 Then Cancel = True
End Sub

Private Sub vsDiagICDEleven_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
     Dim i As Long
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String
    Dim vbPoint As POINTAPI
    Dim blnCancel As Boolean

    Call CreatePublicAdvice
    With vsDiagICDEleven
        Select Case Col
            Case Eleven_诊断描述
                If .TextMatrix(Row, Eleven_诊断类型) = "主干码" Then
                    If gobjPublicAdvice Is Nothing Then Exit Sub
                    Set rsTmp = gobjPublicAdvice.ShowILLSelect(Me, "E", mlng科室ID, , False, False, , , mlngPatiType, 1, True, mintDiagType)
                Else
                    If gobjPublicAdvice Is Nothing Then Exit Sub
                    Set rsTmp = gobjPublicAdvice.ShowILLSelect(Me, "E", mlng科室ID, , False, False, , , mlngPatiType, 1, False, mintDiagType)
                End If
                Call SetICDElvenInput(rsTmp, Row, Col)
        End Select
    End With
End Sub

Private Sub SetICDElvenInput(ByVal rsTmp As ADODB.Recordset, ByVal LngRow As Long, ByVal LngCol As Long)
    Dim i As Long
    With vsDiagICDEleven
        If rsTmp Is Nothing Then Exit Sub
        If rsTmp.EOF Then
            .EditText = .TextMatrix(LngRow, Eleven_诊断描述)
        Else
            For i = 0 To rsTmp.RecordCount - 1
                .TextMatrix(LngRow, Eleven_诊断描述) = rsTmp!名称 & ""
                .EditText = .TextMatrix(LngRow, Eleven_诊断描述)
                .Cell(flexcpData, LngRow, Eleven_诊断描述) = .TextMatrix(LngRow, Eleven_诊断描述)
                .TextMatrix(LngRow, Eleven_诊断编码) = rsTmp!编码 & ""
                .TextMatrix(LngRow, Eleven_疾病ID) = rsTmp!疾病id & ""
                .TextMatrix(LngRow, Eleven_是否病人) = "" & rsTmp!是否病人
                .RowData(LngRow) = rsTmp!项目ID & ""
                rsTmp.MoveNext
            Next
        End If
    End With
End Sub

Private Sub vsDiagICDEleven_DblClick()
    Call vsDiagICDEleven_KeyPress(vbKeySpace)
End Sub

Private Sub vsDiagICDEleven_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        Call imgDelete_Click
    Else
        vsDiagICDEleven_KeyPress KeyCode
    End If
End Sub

Private Sub vsDiagICDEleven_KeyPress(KeyAscii As Integer)
    Dim LngRow As Long, LngCol As Long
    With vsDiagICDEleven
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call EnterCellDiag(.Row, .Col)
        Else
            If Not ICDElevenEditable(.Row, .Col) Then Exit Sub
            Select Case .Col
                Case Eleven_诊断描述
                    If KeyAscii = Asc("*") Then
                        KeyAscii = 0
                        Call vsDiagICDEleven_CellButtonClick(.Row, .Col)
                    Else
                        LngRow = .Row
                        LngCol = .Col
                        .ComboList = ""
                    End If
            End Select
        End If
    End With
End Sub

Private Sub EnterCellDiag(ByVal LngRow As Long, ByVal LngCol As Long)
    Dim i As Long, j As Long

    With vsDiagICDEleven
        '从下一单元开始循环搜索
        If LngRow < .FixedRows Then LngRow = .FixedRows
        For i = LngRow To .Rows - 1
            For j = IIf(i = LngRow, LngCol + 1, Eleven_诊断编码) To Eleven_诊断描述
                If Not .ColHidden(j) Then
                    If ICDElevenEditable(i, j) And .ColWidth(j) <> 0 Then Exit For
                End If
            Next
            If j <= Eleven_诊断描述 Then Exit For
        Next
        If i <= .Rows - 1 Then
            .Row = i: .Col = j
        ElseIf i = .Rows And j > Eleven_诊断描述 And .TextMatrix(.Rows - 1, Eleven_诊断描述) <> "" Then
            If .TextMatrix(.Row, Eleven_诊断类型) = "扩展码" Or .TextMatrix(.Row, Eleven_诊断类型) = "证  候" Then
                If .TextMatrix(.Row - 1, Eleven_诊断描述) <> "" Then
                    .Rows = .Rows + 1
                    Call ChangeVSHeight
                    .TextMatrix(.Rows - 1, Eleven_诊断类型) = .TextMatrix(.Rows - 2, Eleven_诊断类型)
                    .Cell(flexcpBackColor, .Rows - 1, .FixedRows, .Rows - 1, .Cols - 1) = .Cell(flexcpBackColor, .Rows - 2, .FixedRows, .Rows - 2, .Cols - 1)
                    .Row = .Rows - 1: .Col = Eleven_诊断描述
                End If
            End If
        Else
            Call zlcommfun.PressKey(vbKeyTab): mblnEnter = True
        End If
    End With
End Sub


Public Function ICDElevenEditable(ByVal LngRow As Long, ByVal LngCol As Long) As Boolean
    Dim blnJudge As Boolean
    Dim i As Long

    With vsDiagICDEleven
        If LngCol <> Eleven_诊断描述 Then Exit Function
        Select Case LngCol
            Case Eleven_诊断描述
                If .TextMatrix(LngRow, Eleven_诊断描述) <> "" Then
'                    For i = .FixedRows To .Rows - 1
'                        If .TextMatrix(i, Eleven_诊断类型) = .TextMatrix(lngRow, Eleven_诊断类型) And .TextMatrix(i, Eleven_诊断描述) = "" And .TextMatrix(lngRow, Eleven_诊断类型) = "主干码" Then
'                            blnJudge = True
'                        End If
'                    Next
'                    If blnJudge Then Exit Function
                Else
                    If LngRow - 1 >= .FixedRows Then
                        If .TextMatrix(LngRow - 1, Eleven_诊断类型) = "主干码" Then
                            If .TextMatrix(LngRow - 1, Eleven_诊断描述) = "" Then blnJudge = True
                            If blnJudge Then Exit Function
                        End If
                    End If
                End If
        End Select
    End With
    ICDElevenEditable = True
End Function

Private Sub vsDiagICDEleven_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String, blnInputCancel As Boolean
    Dim strInput As String, vPoint As POINTAPI
    Dim strDiagType As String
    Dim strTag As String
    Dim strTmp As String
    Dim i As Long
    Dim LngRow As Long, LngCol As Long

    With vsDiagICDEleven
        LngRow = Row: LngCol = Col
        Select Case LngCol
            Case Eleven_诊断描述
                strTmp = .TextMatrix(LngRow, Eleven_诊断描述)
                If .TextMatrix(LngRow, Eleven_诊断类型) = "主干码" Then
                    strDiagType = IIf(mDiagTag = "西医", decode(mintDiagType, "6", "',2,'", "7", "',23,'", "',1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,27,'"), "',26,'")
                Else
                    strDiagType = IIf(mDiagTag = "西医", "',28,'", "',26,'")
                End If
                If .EditText = "" And .TextMatrix(.Row, .Col) <> "" Then
                    .EditText = ""
                ElseIf .EditText = .Cell(flexcpData, LngRow, LngCol) Then
                    If mblnEnter Then Call EnterCellDiag(LngRow, LngCol)
                ElseIf .TextMatrix(LngRow, Eleven_诊断编码) <> "" And .Cell(flexcpData, LngRow, LngCol) <> "" Then
                    strInput = UCase(.EditText)
                    strSql = GetICDElevenSql(strInput, mstr性别, IIf(.TextMatrix(LngRow, Eleven_诊断类型) = "主干码", 0, 1), strDiagType)
                    vPoint = GetCoordPos(.hwnd, .Left + 15, .CellTop)
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSql, 0, strTag, _
                        False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnInputCancel, False, True, _
                        strInput & "%", mstrLike & strInput & "%", "E", mstr性别, mint简码 + 1, strInput, UserInfo.ID, mlng科室ID)
                    If blnInputCancel Then
                        Cancel = True
                        .EditText = strTmp
                        .TextMatrix(LngRow, Eleven_诊断描述) = strTmp
                        .Cell(flexcpData, LngRow, LngCol) = .TextMatrix(LngRow, Eleven_诊断描述)
                    Else
                        If rsTmp Is Nothing Then
                             Cancel = True
                            .EditText = strTmp
                            .TextMatrix(LngRow, Eleven_诊断描述) = strTmp
                            .Cell(flexcpData, LngRow, LngCol) = .TextMatrix(LngRow, Eleven_诊断描述)
                        Else
                             Call SetICDElvenInput(rsTmp, LngRow, LngCol)
                        End If
                    End If
                Else
                    strInput = UCase(.EditText)
                    strSql = GetICDElevenSql(strInput, mstr性别, IIf(.TextMatrix(LngRow, Eleven_诊断类型) = "主干码", 0, 1), strDiagType)
                    vPoint = GetCoordPos(.hwnd, .Left + 15, .CellTop)
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSql, 0, strTag, _
                        False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnInputCancel, False, True, _
                        strInput & "%", mstrLike & strInput & "%", "E", mstr性别, mint简码 + 1, strInput, UserInfo.ID, mlng科室ID)
                    If blnInputCancel Then
                        Cancel = True
                        .EditText = strTmp
                        .TextMatrix(LngRow, Eleven_诊断描述) = strTmp
                        .Cell(flexcpData, LngRow, LngCol) = .TextMatrix(LngRow, Eleven_诊断描述)
                    Else
                        If Not rsTmp Is Nothing Then
                            Call SetICDElvenInput(rsTmp, LngRow, LngCol)
                        Else
                            Cancel = True
                            .EditText = strTmp
                            .TextMatrix(LngRow, Eleven_诊断描述) = strTmp
                            .Cell(flexcpData, LngRow, LngCol) = .TextMatrix(LngRow, Eleven_诊断描述)
                        End If
                    End If
                End If
        End Select
    End With
End Sub

Private Function GetICDElevenSql(ByVal strInput As String, ByRef str性别 As String, ByVal intType As Integer, Optional ByVal strOtherInfo As String) As String
    Dim strSql As String
    Dim lng疾病序号 As Long, lng证候序号 As Long
    Dim rsTmp As ADODB.Recordset
    
    If strOtherInfo = "',26,'" Then
        strSql = "Select 序号 From 疾病编码分类 Where 章节 = '26' And 名称 = '传统医学疾病（TM1）' And 编码 = 'L1-SA0'"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "疾病编码分类")
        If Not rsTmp.EOF Then
            lng疾病序号 = Val("" & rsTmp!序号)
        End If
        
        strSql = "Select 序号 From 疾病编码分类 Where 章节 = '26' And 名称 = '传统医学证候（TM1）' And 编码 = 'L1-SE7'"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "疾病编码分类")
        If Not rsTmp.EOF Then
            lng证候序号 = Val("" & rsTmp!序号)
        End If
    End If
    If zlcommfun.IsCharChinese(strInput) Then
        strSql = "A.名称 Like [2]" & IIf(intType = 0, IIf(lng疾病序号 <> 0 And lng证候序号 <> 0, " And C.序号>= " & lng疾病序号 & " And C.序号<" & lng证候序号, ""), IIf(lng证候序号 <> 0, " And C.序号>" & lng证候序号, "")) '输入汉字时只匹配名称
    Else
        strSql = "A.编码 Like [1] Or A.名称 Like [2] Or " & IIf(mint简码 = 0, "A.简码", "A.五笔码") & " Like [2]" & IIf(intType = 0, IIf(lng疾病序号 <> 0 And lng证候序号 <> 0, " And C.序号>= " & lng疾病序号 & " And C.序号<" & lng证候序号, ""), IIf(lng证候序号 <> 0, " And C.序号>" & lng证候序号, ""))
    End If
    strSql = "Select A.Id, A.Id 项目ID,A.编码,A.名称," & IIf(mint简码 = 0, "A.简码", "A.五笔码") & " as 简码,  C.是否病人,A.编码 疾病编码, A.Id 疾病id,A.类别 疾病类别" & vbNewLine & _
        "From 疾病编码目录 A, 疾病编码分类 C" & vbNewLine & _
        "Where A.分类id = C.Id(+) And A.章节=C.章节(+) " & IIf(strOtherInfo <> "", " And Instr(" & strOtherInfo & "," & " ',' || A.章节 || ',')>0", "") & " And Instr([3],A.类别)>0 And (" & strSql & ")" & _
        IIf(str性别 <> "", " And (A.性别限制=[4] Or A.性别限制 is NULL)", "") & _
        IIf(mlngPatiType = 1, " And (Nvl(A.适用范围,0) = 0 or A.适用范围 =1) ", " And (Nvl(A.适用范围,0) = 0 or A.适用范围 =2) ") & vbNewLine & _
        " And (A.撤档时间 is Null Or A.撤档时间=To_Date('3000-01-01','YYYY-MM-DD'))" & _
        " Order by A.编码"
    strSql = "Select distinct A.Id,A.项目ID, A.编码, A.名称,A.简码, A.是否病人,A.疾病编码, A.疾病id,A.疾病类别, " & _
                " Decode(a.名称, [6], 1, Decode(A.简码,[6],1,decode(a.编码,[6],1,NULL))) As 排序1ID," & vbNewLine & _
        "                Decode(d.疾病id, Null, Decode(c.疾病id, Null, Null, 2), 1) As 排序2ID," & vbNewLine & _
        "                Decode(Substr(a.名称, 1, Length([6])), [6], 1, Decode(Substr(A.简码, 1, Length([6])),[6],1,decode(Substr(a.编码, 1, Length([6])),[6],1,NULL))) As 排序3ID" & vbNewLine & _
                " From (" & strSql & ") A, 疾病编码科室 C, 疾病编码科室 D " & _
                " Where  c.疾病id(+) = a.Id And d.疾病id(+) = a.Id And c.科室id(+)=[8]  And d.人员id(+) = [7] " & _
                " Order By 疾病类别 desc,排序1ID, 排序2ID, 排序3ID, A.编码"
    GetICDElevenSql = strSql
End Function

Private Function MoveCurrRow(ByVal strType As String, ByVal LngRow As Long, ByVal lngWay As Long) As Long
'功能：将当前行上移或下移一行
'参数：lngRow=当前行
'      lngWay=1上移一行,-1下移一行(相当于下一行上移一行)
    Dim lngPreRow As Long, lngNextRow As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim lngUpBegin As Long, lngUpEnd As Long
    Dim lngDownBegin As Long, lngDownEnd As Long
    Dim i As Long, j As Long
    Dim lngMoveRows As Long, blnRedraw As Boolean
    With vsDiagICDEleven
        Call GetRowScope(strType, LngRow, lngBegin, lngEnd)
        If lngWay = 1 Then
            lngPreRow = GetPreRow(lngBegin)
            If lngPreRow = -1 Then Exit Function
            lngDownBegin = lngBegin
            lngDownEnd = lngEnd
            Call GetRowScope(strType, lngPreRow, lngUpBegin, lngUpEnd)
            lngMoveRows = lngDownBegin - lngUpBegin
        Else
            lngNextRow = GetNextRow(lngEnd)
            If lngNextRow = -1 Then Exit Function
            lngUpBegin = lngBegin
            lngUpEnd = lngEnd
            Call GetRowScope(strType, lngNextRow, lngDownBegin, lngDownEnd)
            lngMoveRows = lngDownEnd - lngUpEnd
        End If

        MoveCurrRow = lngMoveRows
        j = 0
        For i = lngDownBegin To lngDownEnd
            .RowPosition(i) = lngUpBegin + j
            j = j + 1
        Next
        
        LngRow = LngRow - lngWay * lngMoveRows
        .Row = LngRow
    End With
End Function

Private Sub GetRowScope(ByVal strType As String, ByVal LngRow As Long, lngBegin As Long, lngEnd As Long)
    Dim i As Long, j As Long, k As Long
    With vsDiagICDEleven
        lngBegin = LngRow: lngEnd = LngRow
        If strType = "扩展码" Or strType = "证  候" Then
            j = LngRow
            For i = LngRow To .FixedRows Step -1
                If strType = .TextMatrix(i, Eleven_诊断类型) Then
                    j = i
                Else
                    Exit For
                End If
            Next
            k = LngRow
            For i = .Row To .Rows - 1
                If strType = .TextMatrix(i, Eleven_诊断类型) Then
                    k = i
                Else
                    Exit For
                End If
            Next
        Else
            j = LngRow
            For i = LngRow To .FixedRows Step -1
                If strType = .TextMatrix(i, Eleven_诊断类型) Then
                    j = i
                    Exit For
                End If
            Next
            k = LngRow
            For i = LngRow To .Rows - 1
                If strType <> .TextMatrix(i, Eleven_诊断类型) Then
                    k = i
                Else
                    If i <> LngRow Then
                        Exit For
                    End If
                End If
            Next
            lngBegin = j: lngEnd = k
        End If
    End With
End Sub

Private Function GetPreRow(ByVal LngRow As Long) As Long
'功能：取上一最近行
'返回：无有效行时,返回-1
    Dim lngTmp As Long, i As Long

    lngTmp = -1
    For i = LngRow - 1 To vsDiagICDEleven.FixedRows Step -1
        lngTmp = i: Exit For
    Next
    GetPreRow = lngTmp
End Function

Private Function GetNextRow(ByVal LngRow As Long) As Long
'功能：取下一最近行
'返回：无有效行时,返回-1
    Dim lngTmp As Long, i As Long

    lngTmp = -1
    For i = LngRow + 1 To vsDiagICDEleven.Rows - 1
        lngTmp = i: Exit For
    Next
    GetNextRow = lngTmp
End Function

Public Sub InitRsICDEleven(ByRef rsData As ADODB.Recordset)
'功能：初始化记录集
    Set rsData = New ADODB.Recordset
    With rsData
        .Fields.Append "行号", adInteger '初始加载时用
        
        .Fields.Append "诊断编码", adVarChar, 2000 '疾病编码目录.编码
        .Fields.Append "诊断描述", adVarChar, 4000 '疾病编码目录.名称
        .Fields.Append "诊断类型", adVarChar, 100 '区分主干码和扩展码，字符串，"主干码"/"扩展码"
        .Fields.Append "IndexEx", adVarChar, 4 ' 录入次序 如 01,0101,0102,0103,02,0201,0202,0203
        .Fields.Append "疾病ID", adInteger, 100 '疾病编码目录.ID
        .Fields.Append "医嘱IDs", adVarChar, 200
        .Fields.Append "Tag", adVarChar, 4000 '中医,西医
        .Fields.Append "标记", adInteger '界面表格上的行序号 .row
        .Fields.Append "信息名", adVarChar, 100 '
        
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .CursorType = adOpenStatic
        .Open
    End With
End Sub

Public Sub GetRsICD11录入主行(ByRef rsData As ADODB.Recordset, ByRef lng疾病ID As Long, ByRef str编码Out As String, ByRef str描述Out As String)
'功能：加工记录集的ICD11缓存数据以适应接口
'参数：intType 0-传入前加工，1-返回后加工
    Dim i As Long
    Dim lng主序 As Long
    Dim lng次序 As Long
    Dim strBackInfo As String
    Dim str编码 As String
    Dim rsTmp As ADODB.Recordset
    Dim lng序号 As Long
           
    With rsData
        .Filter = 0
        .Sort = "IndexEx"
        lng疾病ID = Val(!疾病id & "")
        For i = 1 To rsData.RecordCount
            If !诊断类型 & "" = "主干码" Then
                strBackInfo = strBackInfo & "/" & !诊断描述
                str编码 = str编码 & "/" & !诊断编码
            Else
                strBackInfo = strBackInfo & "&" & !诊断描述
                str编码 = str编码 & "&" & !诊断编码
            End If
            .MoveNext
        Next
        str编码Out = Mid(str编码, 2)
        str描述Out = Mid(strBackInfo, 2)
        .Filter = 0
        .Sort = "IndexEx"
    End With
End Sub











