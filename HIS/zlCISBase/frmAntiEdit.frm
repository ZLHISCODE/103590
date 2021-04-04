VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAntiEdit 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7065
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10425
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7065
   ScaleWidth      =   10425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picName 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1395
      Left            =   0
      ScaleHeight     =   1395
      ScaleWidth      =   5370
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   5370
      Begin VB.TextBox txt英文 
         Height          =   300
         Left            =   615
         MaxLength       =   60
         TabIndex        =   8
         Top             =   960
         Width           =   4545
      End
      Begin VB.TextBox txt中文 
         Height          =   300
         Left            =   615
         MaxLength       =   60
         TabIndex        =   6
         Top             =   540
         Width           =   4545
      End
      Begin VB.TextBox txt编码 
         Height          =   300
         Left            =   615
         MaxLength       =   10
         TabIndex        =   2
         Top             =   120
         Width           =   1635
      End
      Begin VB.TextBox txt缩写 
         Height          =   300
         Left            =   3525
         MaxLength       =   10
         TabIndex        =   4
         Top             =   120
         Width           =   1635
      End
      Begin VB.Label lbl英文 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "英文"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   165
         TabIndex        =   7
         Top             =   1020
         Width           =   360
      End
      Begin VB.Label lbl中文 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "中文"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   165
         TabIndex        =   5
         Top             =   615
         Width           =   360
      End
      Begin VB.Label lbl编码 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "编码"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   165
         TabIndex        =   1
         Top             =   195
         Width           =   360
      End
      Begin VB.Label lbl缩写 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "缩写"
         Height          =   180
         Left            =   3075
         TabIndex        =   3
         Top             =   180
         Width           =   360
      End
   End
   Begin VB.PictureBox picSingle 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   4275
      Left            =   5000
      ScaleHeight     =   4275
      ScaleWidth      =   5370
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1395
      Visible         =   0   'False
      Width           =   5370
      Begin VB.ComboBox cbo药敏方法 
         Height          =   300
         Left            =   1335
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   150
         Width           =   1335
      End
      Begin VB.TextBox txtWHONET码 
         Height          =   300
         Left            =   3900
         MaxLength       =   10
         TabIndex        =   13
         Top             =   150
         Width           =   1260
      End
      Begin VB.TextBox txt说明 
         Height          =   300
         Left            =   1335
         MaxLength       =   10
         TabIndex        =   15
         Top             =   570
         Width           =   3825
      End
      Begin VB.Frame frmLine 
         Height          =   15
         Left            =   165
         TabIndex        =   40
         Top             =   0
         Width           =   5040
      End
      Begin VB.Frame fraLine 
         Height          =   15
         Index           =   1
         Left            =   165
         TabIndex        =   39
         Top             =   1005
         Width           =   5055
      End
      Begin VB.TextBox txt用法用量 
         Height          =   300
         Index           =   0
         Left            =   1170
         MaxLength       =   10
         TabIndex        =   17
         Top             =   1155
         Width           =   3990
      End
      Begin VB.TextBox txt血药浓度 
         Height          =   300
         Index           =   0
         Left            =   1965
         MaxLength       =   10
         TabIndex        =   19
         Top             =   1560
         Width           =   3195
      End
      Begin VB.TextBox txt尿药浓度 
         Height          =   300
         Index           =   0
         Left            =   1965
         MaxLength       =   10
         TabIndex        =   21
         Top             =   1965
         Width           =   3195
      End
      Begin VB.TextBox txt用法用量 
         Height          =   300
         Index           =   1
         Left            =   1170
         MaxLength       =   10
         TabIndex        =   23
         Top             =   2400
         Width           =   3990
      End
      Begin VB.TextBox txt血药浓度 
         Height          =   300
         Index           =   1
         Left            =   1965
         MaxLength       =   10
         TabIndex        =   25
         Top             =   2805
         Width           =   3195
      End
      Begin VB.TextBox txt尿药浓度 
         Height          =   300
         Index           =   1
         Left            =   1965
         MaxLength       =   10
         TabIndex        =   27
         Top             =   3210
         Width           =   3195
      End
      Begin VB.Frame fraLine 
         Height          =   15
         Index           =   0
         Left            =   165
         TabIndex        =   38
         Top             =   3660
         Width           =   5055
      End
      Begin VB.Label lblNote 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "用法用量主要用于微生物检验试验报告中供临床的参考。"
         Height          =   180
         Left            =   555
         TabIndex        =   28
         Top             =   3810
         Width           =   4500
      End
      Begin VB.Image imgNote 
         Height          =   240
         Left            =   180
         Picture         =   "frmAntiEdit.frx":0000
         Top             =   3780
         Width           =   240
      End
      Begin VB.Label lbl药敏方法 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "默认药敏方法"
         Height          =   180
         Left            =   165
         TabIndex        =   10
         Top             =   210
         Width           =   1080
      End
      Begin VB.Label lblWHONET码 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "WHONET码"
         Height          =   180
         Left            =   3135
         TabIndex        =   12
         Top             =   210
         Width           =   720
      End
      Begin VB.Label lbl说明 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "药品附加说明"
         Height          =   180
         Left            =   165
         TabIndex        =   14
         Top             =   630
         Width           =   1080
      End
      Begin VB.Label lbl用法用量 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "用法用量①"
         Height          =   180
         Index           =   0
         Left            =   165
         TabIndex        =   16
         Top             =   1215
         Width           =   900
      End
      Begin VB.Label lbl血药浓度 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "血药浓度"
         Height          =   180
         Index           =   0
         Left            =   1170
         TabIndex        =   18
         Top             =   1620
         Width           =   720
      End
      Begin VB.Label lbl尿药浓度 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "尿药浓度"
         Height          =   180
         Index           =   0
         Left            =   1170
         TabIndex        =   20
         Top             =   2025
         Width           =   720
      End
      Begin VB.Label lbl用法用量 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "用法用量②"
         Height          =   180
         Index           =   1
         Left            =   165
         TabIndex        =   22
         Top             =   2460
         Width           =   900
      End
      Begin VB.Label lbl血药浓度 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "血药浓度"
         Height          =   180
         Index           =   1
         Left            =   1170
         TabIndex        =   24
         Top             =   2865
         Width           =   720
      End
      Begin VB.Label lbl尿药浓度 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "尿药浓度"
         Height          =   180
         Index           =   1
         Left            =   1170
         TabIndex        =   26
         Top             =   3270
         Width           =   720
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgGroup 
      Height          =   2805
      Left            =   165
      TabIndex        =   30
      Top             =   1665
      Visible         =   0   'False
      Width           =   5040
      _cx             =   8890
      _cy             =   4948
      Appearance      =   2
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
      Cols            =   3
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
   Begin VB.PictureBox picGroup 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   2505
      Left            =   165
      ScaleHeight     =   2505
      ScaleWidth      =   5040
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   4530
      Visible         =   0   'False
      Width           =   5040
      Begin VB.TextBox txtFind 
         Height          =   300
         Left            =   540
         TabIndex        =   33
         Top             =   70
         Width           =   1650
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "…"
         Height          =   315
         Left            =   2190
         TabIndex        =   34
         TabStop         =   0   'False
         ToolTipText     =   "查找符合条件的项目"
         Top             =   63
         Width           =   360
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "∧ 添加"
         Height          =   350
         Index           =   0
         Left            =   2925
         TabIndex        =   35
         Top             =   45
         Width           =   1080
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "∨ 删除"
         Height          =   350
         Index           =   1
         Left            =   4005
         TabIndex        =   36
         Top             =   45
         Width           =   1080
      End
      Begin MSComctlLib.ListView lvwGroup 
         Height          =   2040
         Left            =   0
         TabIndex        =   37
         Top             =   450
         Width           =   5025
         _ExtentX        =   8864
         _ExtentY        =   3598
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label lblFind 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "查找:"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   75
         TabIndex        =   32
         Top             =   130
         Width           =   450
      End
   End
   Begin VB.Label lblGroup 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "该药敏组包含的抗生素:"
      Height          =   180
      Left            =   165
      TabIndex        =   29
      Top             =   1440
      Width           =   1890
   End
End
Attribute VB_Name = "frmAntiEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngItemID As Long          '当前显示的项目id
Private mintGroup As Integer        '当前显示的项目id

Private Enum mcol
    ID = 0: 编码: 中文名: 缩写
End Enum

Dim objItem As ListItem
Dim lngCount As Long

'--------------------------------------------
'以下为窗体公共方法
'--------------------------------------------
Private Sub setListFormat(Optional blnKeepData As Boolean)
    '功能：初始化设置列表
    '参数： blnKeepData-是否保留数据，即只是重新设置格式
    With Me.vfgGroup
        .Redraw = flexRDNone
        If blnKeepData = False Then
            .Clear
            .Rows = 1: .FixedRows = 1: .Cols = 4: .FixedCols = 0
        End If
        .TextMatrix(0, mcol.ID) = "ID": .TextMatrix(0, mcol.编码) = "编码"
        .TextMatrix(0, mcol.中文名) = "中文名": .TextMatrix(0, mcol.缩写) = "缩写"
        
        .ColWidth(mcol.ID) = 0:  .ColWidth(mcol.编码) = 900
        .ColWidth(mcol.中文名) = 3000: .ColWidth(mcol.编码) = 1000
        For lngCount = 0 To .Cols - 1
            .FixedAlignment(lngCount) = flexAlignCenterCenter
            If .ColWidth(lngCount) = 0 Then .ColHidden(lngCount) = True
        Next
        .Redraw = flexRDDirect
    End With
End Sub

Public Function zlRefresh(lngItemID As Long, intGroup As Integer) As Boolean
    '功能：根据项目id刷新当前显示内容
    Dim rsTemp As New ADODB.Recordset
    
    mlngItemID = lngItemID: mintGroup = intGroup
    
    '清除此前项目的显示
    Me.txt编码.Text = "": Me.txt中文.Text = "": Me.txt英文.Text = "": Me.txt缩写.Text = ""
    If intGroup = 0 Then
        Me.txtWHONET码.Text = "": Me.txt说明.Text = ""
        Me.txt用法用量(0).Text = "": Me.txt血药浓度(0).Text = "": Me.txt尿药浓度(0).Text = ""
        Me.txt用法用量(1).Text = "": Me.txt血药浓度(1).Text = "": Me.txt尿药浓度(1).Text = ""
        
        Me.picSingle.Visible = True: Me.vfgGroup.Visible = False: Me.picGroup.Visible = False
    Else
        Me.txtFind.Text = "": Me.lvwGroup.ListItems.Clear: Call setListFormat
        Me.picSingle.Visible = False: Me.vfgGroup.Visible = True: Me.picGroup.Visible = True
    End If
    If lngItemID = 0 Then zlRefresh = True: Exit Function
    
    '获取指定项目的信息
    Err = 0: On Error GoTo ErrHand
    
    If intGroup = 0 Then
        gstrSql = "Select 编码, 中文名, 英文名, 简码, 说明, 药敏方法, Whonet码, 用法用量1, 血药浓度1, 尿药浓度1, 用法用量2, 血药浓度2," & vbNewLine & _
                "       尿药浓度2" & vbNewLine & _
                "From 检验用抗生素" & vbNewLine & _
                "Where ID = [1]"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemID)
        With rsTemp
            Me.txt编码.MaxLength = .Fields("编码").DefinedSize: Me.txt中文.MaxLength = .Fields("中文名").DefinedSize
            Me.txt英文.MaxLength = .Fields("英文名").DefinedSize: Me.txt缩写.MaxLength = .Fields("简码").DefinedSize
            Me.txt说明.MaxLength = .Fields("说明").DefinedSize: Me.txtWHONET码.MaxLength = .Fields("WHONET码").DefinedSize
            Me.txt用法用量(0).MaxLength = .Fields("用法用量1").DefinedSize: Me.txt用法用量(1).MaxLength = Me.txt用法用量(0).MaxLength
            Me.txt血药浓度(0).MaxLength = .Fields("血药浓度1").DefinedSize: Me.txt血药浓度(1).MaxLength = Me.txt血药浓度(0).MaxLength
            Me.txt尿药浓度(0).MaxLength = .Fields("尿药浓度1").DefinedSize: Me.txt尿药浓度(1).MaxLength = Me.txt尿药浓度(0).MaxLength
            If .RecordCount > 0 Then
                Me.txt编码.Text = "" & !编码: Me.txt中文.Text = "" & !中文名
                Me.txt英文.Text = "" & !英文名: Me.txt缩写.Text = "" & !简码
                Me.txt说明.Text = "" & !说明: Me.txtWHONET码.Text = "" & !WHONET码
                If Val("" & !药敏方法) = 3 Then
                    Me.cbo药敏方法.ListIndex = 2
                ElseIf Val("" & !药敏方法) = 2 Then
                    Me.cbo药敏方法.ListIndex = 1
                Else
                    Me.cbo药敏方法.ListIndex = 0
                End If
                Me.txt用法用量(0).Text = "" & !用法用量1: Me.txt用法用量(1).Text = "" & !用法用量2
                Me.txt血药浓度(0).Text = "" & !血药浓度1: Me.txt血药浓度(1).Text = "" & !血药浓度2
                Me.txt尿药浓度(0).Text = "" & !尿药浓度1: Me.txt尿药浓度(1).Text = "" & !尿药浓度2
            End If
        End With
    Else
        gstrSql = "Select 编码, 名称, 英文, 简码 From 检验抗生素组 Where ID = [1]"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemID)
        With rsTemp
            Me.txt编码.MaxLength = .Fields("编码").DefinedSize: Me.txt中文.MaxLength = .Fields("名称").DefinedSize
            Me.txt英文.MaxLength = .Fields("英文").DefinedSize: Me.txt缩写.MaxLength = .Fields("简码").DefinedSize
            If .RecordCount > 0 Then
                Me.txt编码.Text = "" & !编码: Me.txt中文.Text = "" & !名称
                Me.txt英文.Text = "" & !英文: Me.txt缩写.Text = "" & !简码
            End If
        End With
        
        gstrSql = "Select I.ID, I.编码, I.中文名, I.简码 As 缩写" & vbNewLine & _
                "From 检验抗生素用药 L, 检验用抗生素 I" & vbNewLine & _
                "Where L.抗生素id = I.ID And L.抗生素分组id = [1]"
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemID)
        Set Me.vfgGroup.DataSource = rsTemp: Call setListFormat(True)
        If Me.vfgGroup.Rows > Me.vfgGroup.FixedRows Then Me.vfgGroup.Row = Me.vfgGroup.FixedRows
    
    End If
    zlRefresh = True: Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlRefresh = False: Exit Function
End Function

Public Function zlEditStart(blnAdd As Boolean, lngItemID As Long, intGroup As Integer) As Boolean
    '功能：开始项目编辑
    '参数： blnAdd-是否增加，否则为修改
    '       lngItemId-增加的参照项目，或者指定编辑的项目
    Dim rsTemp As New ADODB.Recordset
    
    mintGroup = intGroup
    
    If blnAdd Then
        Err = 0: On Error GoTo ErrHand
        If intGroup = 0 Then
            gstrSql = "Select Nvl(Max(To_Number(编码)), 0) As 编码, Nvl(Max(Length(编码)), 0) As 长度 From 检验用抗生素"
        Else
            gstrSql = "Select Nvl(Max(To_Number(编码)), 0) As 编码, Nvl(Max(Length(编码)), 0) As 长度 From 检验抗生素组"
        End If
        
        Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, "cmd产地_Click")
        With rsTemp
'            If .State = adStateOpen Then .Close
'            Call SQLTest(App.ProductName, Me.Caption, gstrSql)
'            Call SQLTest
            If !长度 <> 0 And !长度 <= Me.txt编码.MaxLength Then
                Me.txt编码.Text = Format(Val(!编码) + 1, String(!长度, "0"))
            Else
                Me.txt编码.Text = Format(Val(!编码) + 1, String(Me.txt编码.MaxLength, "0"))
            End If
        End With
        Me.txt中文.Text = "": Me.txt英文.Text = "": Me.txt缩写.Text = ""
        If intGroup = 0 Then
            Me.txtWHONET码.Text = "": Me.txt说明.Text = ""
            Me.txt用法用量(0).Text = "": Me.txt血药浓度(0).Text = "": Me.txt尿药浓度(0).Text = ""
            Me.txt用法用量(1).Text = "": Me.txt血药浓度(1).Text = "": Me.txt尿药浓度(1).Text = ""
            Me.picSingle.Visible = True: Me.vfgGroup.Visible = False: Me.picGroup.Visible = False
        Else
            Me.txtFind.Text = "": Me.lvwGroup.ListItems.Clear: Call setListFormat
            Me.picSingle.Visible = False: Me.vfgGroup.Visible = True: Me.picGroup.Visible = True
        End If
    End If

    Me.Tag = IIf(blnAdd, "增加", "修改")
    Me.BackColor = RGB(250, 250, 250)
    Me.picName.BackColor = Me.BackColor: Me.picSingle.BackColor = Me.BackColor: Me.picGroup.BackColor = Me.BackColor
    Me.picName.Enabled = True
    If intGroup = 0 Then
        Me.picSingle.Enabled = True
    Else
        Me.picGroup.Enabled = True
        Call Form_Resize
    End If
    
    Me.txt编码.SetFocus
    zlEditStart = True: Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlEditStart = False: Exit Function
End Function

Public Sub zlEditCancel()
    '功能：放弃正在进行的编辑
    Me.Tag = ""
    Me.BackColor = &H8000000F
    Me.picName.BackColor = Me.BackColor: Me.picSingle.BackColor = Me.BackColor: Me.picGroup.BackColor = Me.BackColor
    Me.picName.Enabled = True
    Me.picSingle.Enabled = True
    Me.picGroup.Enabled = True
    Call Form_Resize
    
'    Call Me.zlRefresh(mlngItemID, mintGroup)
End Sub

Public Function zlEditSave() As Long
    '功能：保存正在进行的编辑,并返回正在编辑项目id,保存失败返回0
    Dim lngNewId As Long, strLists As String
    
    '一般特性检查
    If Trim(Me.txt编码.Text) = "" Then
        MsgBox "请输入编码！", vbInformation, gstrSysName
        Me.txt编码.SetFocus: zlEditSave = 0: Exit Function
    End If
    If Val(Me.txt编码.Text) > Val(String(Me.txt编码.MaxLength, "9")) Then
        MsgBox "编码太大！", vbInformation, gstrSysName
        Me.txt编码.SetFocus: zlEditSave = 0: Exit Function
    End If
    If Trim(Me.txt中文.Text) = "" Then
        MsgBox "请输入中文名称！", vbInformation, gstrSysName
        Me.txt中文.SetFocus: zlEditSave = 0: Exit Function
    End If
    If LenB(StrConv(Trim(Me.txt中文.Text), vbFromUnicode)) > Me.txt中文.MaxLength Then
        MsgBox "中文名称超长（最多" & Me.txt中文.MaxLength & "个字符或等长汉字）！", vbInformation, gstrSysName
        Me.txt中文.SetFocus: zlEditSave = 0: Exit Function
    End If
    If LenB(StrConv(Trim(Me.txt英文.Text), vbFromUnicode)) > Me.txt英文.MaxLength Then
        MsgBox "英文名称超长（最多" & Me.txt英文.MaxLength & "个字符）！", vbInformation, gstrSysName
        Me.txt英文.SetFocus: zlEditSave = 0: Exit Function
    End If
    If LenB(StrConv(Trim(Me.txt缩写.Text), vbFromUnicode)) > Me.txt缩写.MaxLength Then
        MsgBox "缩写超长（最多" & Me.txt缩写.MaxLength & "个字符）！", vbInformation, gstrSysName
        Me.txt缩写.SetFocus: zlEditSave = 0: Exit Function
    End If
    
    gstrSql = "'" & Trim(Me.txt编码.Text) & "','" & Trim(Me.txt中文.Text) & "'"
    gstrSql = gstrSql & ",'" & Trim(Me.txt英文.Text) & "','" & Trim(Me.txt缩写.Text) & "'"
    If mintGroup = 0 Then
        If Me.cbo药敏方法.ListIndex = -1 Then Me.cbo药敏方法.ListIndex = 0
        If LenB(StrConv(Trim(Me.txtWHONET码.Text), vbFromUnicode)) > Me.txtWHONET码.MaxLength Then
            MsgBox "WHONET码超长（最多" & Me.txtWHONET码.MaxLength & "个字符）！", vbInformation, gstrSysName
            Me.txtWHONET码.SetFocus: zlEditSave = 0: Exit Function
        End If
        If LenB(StrConv(Trim(Me.txt说明.Text), vbFromUnicode)) > Me.txt说明.MaxLength Then
            MsgBox "附加说明超长（最多" & Me.txt说明.MaxLength & "个字符）！", vbInformation, gstrSysName
            Me.txt说明.SetFocus: zlEditSave = 0: Exit Function
        End If
        For lngCount = 0 To 1
            If LenB(StrConv(Trim(Me.txt用法用量(lngCount).Text), vbFromUnicode)) > Me.txt用法用量(lngCount).MaxLength Then
                MsgBox "用法用量" & "lngCount" & "超长（最多" & Me.txt用法用量(lngCount).MaxLength & "个字符）！", vbInformation, gstrSysName
                Me.txt用法用量(lngCount).SetFocus: zlEditSave = 0: Exit Function
            End If
            If LenB(StrConv(Trim(Me.txt血药浓度(lngCount).Text), vbFromUnicode)) > Me.txt血药浓度(lngCount).MaxLength Then
                MsgBox "血药浓度" & "lngCount" & "超长（最多" & Me.txt血药浓度(lngCount).MaxLength & "个字符）！", vbInformation, gstrSysName
                Me.txt血药浓度(lngCount).SetFocus: zlEditSave = 0: Exit Function
            End If
            If LenB(StrConv(Trim(Me.txt尿药浓度(lngCount).Text), vbFromUnicode)) > Me.txt尿药浓度(lngCount).MaxLength Then
                MsgBox "尿药浓度" & "lngCount" & "超长（最多" & Me.txt尿药浓度(lngCount).MaxLength & "个字符）！", vbInformation, gstrSysName
                Me.txt尿药浓度(lngCount).SetFocus: zlEditSave = 0: Exit Function
            End If
        Next
    
        gstrSql = gstrSql & ",'" & Trim(Me.txt说明.Text) & "'," & Me.cbo药敏方法.ListIndex + 1 & ",'" & Trim(Me.txtWHONET码.Text) & "'"
        gstrSql = gstrSql & ",'" & Trim(Me.txt用法用量(0).Text) & "','" & Trim(Me.txt血药浓度(0).Text) & "','" & Trim(Me.txt尿药浓度(0).Text) & "'"
        gstrSql = gstrSql & ",'" & Trim(Me.txt用法用量(1).Text) & "','" & Trim(Me.txt血药浓度(1).Text) & "','" & Trim(Me.txt尿药浓度(1).Text) & "'"
    
    Else
        strLists = ""
        With Me.vfgGroup
            For lngCount = .FixedRows To .Rows - 1
                strLists = strLists & "," & .TextMatrix(lngCount, mcol.ID)
            Next
        End With
        If strLists <> "" Then strLists = Mid(strLists, 2)
        gstrSql = gstrSql & ",'" & strLists & "'"
    End If
    
    '数据保存语句组织
    
    lngNewId = mlngItemID
    If mintGroup = 0 Then
        If Me.Tag = "增加" Then
            lngNewId = zldatabase.GetNextId("检验用抗生素")
            gstrSql = "Zl_检验用抗生素_Edit(1," & lngNewId & "," & gstrSql & ")"
        Else
            gstrSql = "Zl_检验用抗生素_Edit(2," & lngNewId & "," & gstrSql & ")"
        End If
    Else
        If Me.Tag = "增加" Then
            lngNewId = zldatabase.GetNextId("检验抗生素组")
            gstrSql = "Zl_检验抗生素组_Edit(1," & lngNewId & "," & gstrSql & ")"
        Else
            gstrSql = "Zl_检验抗生素组_Edit(2," & lngNewId & "," & gstrSql & ")"
        End If
    End If
    
    Err = 0: On Error GoTo ErrHand
    Call SQLTest(App.ProductName, Me.Caption, gstrSql): gcnOracle.Execute gstrSql, , adCmdStoredProc: Call SQLTest
    
    If Me.Tag = "增加" Then mlngItemID = lngNewId
    
    Me.Tag = ""
    Me.BackColor = &H8000000F
    Me.picName.BackColor = Me.BackColor: Me.picSingle.BackColor = Me.BackColor: Me.picGroup.BackColor = Me.BackColor
    Me.picName.Enabled = True
    Me.picSingle.Enabled = True
    Me.picGroup.Enabled = True
    Call Form_Resize
    
    zlEditSave = mlngItemID: Exit Function
    
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlEditSave = 0: Exit Function
End Function

'--------------------------------------------
'以下为窗体控件响应事件
'--------------------------------------------

Private Sub cbo药敏方法_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cmdEdit_Click(Index As Integer)
    Dim lngCurRow As Long
    With Me.vfgGroup
        Select Case Index
        Case 0         '添加
            If Me.lvwGroup.SelectedItem Is Nothing Then Exit Sub
            Set objItem = Me.lvwGroup.SelectedItem
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, mcol.ID) = Mid(objItem.Key, 2)
            .TextMatrix(.Rows - 1, mcol.编码) = objItem.Text
            .TextMatrix(.Rows - 1, mcol.中文名) = objItem.SubItems(Me.lvwGroup.ColumnHeaders("_中文名").Index - 1)
            .TextMatrix(.Rows - 1, mcol.缩写) = objItem.SubItems(Me.lvwGroup.ColumnHeaders("_缩写").Index - 1)
            If .Rows > .FixedRows And .Row < .FixedRows Then .Row = .FixedRows
            Me.lvwGroup.ListItems.Remove objItem.Key: Me.lvwGroup.SetFocus
        Case 1          '删除
            If .Row < .FixedRows Then Exit Sub
            Set objItem = Me.lvwGroup.ListItems.Add(, "_" & .TextMatrix(.Row, mcol.ID), .TextMatrix(.Row, mcol.编码))
            objItem.SubItems(Me.lvwGroup.ColumnHeaders("_中文名").Index - 1) = .TextMatrix(.Row, mcol.中文名)
            objItem.SubItems(Me.lvwGroup.ColumnHeaders("_缩写").Index - 1) = .TextMatrix(.Row, mcol.缩写)
            objItem.Selected = True
            .RemoveItem .Row
        End Select
        .SetFocus
    End With
End Sub

Private Sub cmdFind_Click()
    Dim rsTemp As New ADODB.Recordset
    Dim strFind As String
    strFind = Trim(UCase(Me.txtFind.Text))
    gstrSql = "Select ID, 编码, 中文名, 简码" & vbNewLine & _
            "From 检验用抗生素" & vbNewLine & _
            "Where 编码 Like '" & strFind & "%' Or Upper(中文名) Like '" & gstrMatch & strFind & "%' Or" & vbNewLine & _
            "      Upper(英文名) Like '" & gstrMatch & strFind & "%' Or Upper(简码) Like '" & gstrMatch & strFind & "%'"
    Err = 0: On Error GoTo ErrHand
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSql, Me.Caption)
    With rsTemp
        Me.lvwGroup.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwGroup.ListItems.Add(, "_" & !ID, !编码)
            objItem.SubItems(Me.lvwGroup.ColumnHeaders("_中文名").Index - 1) = "" & !中文名
            objItem.SubItems(Me.lvwGroup.ColumnHeaders("_缩写").Index - 1) = "" & !简码
            .MoveNext
        Loop
    End With
    
    Err = 0: On Error Resume Next
    With Me.vfgGroup
        For lngCount = .FixedRows To .Rows - 1
            Me.lvwGroup.ListItems.Remove "_" & .TextMatrix(lngCount, mcol.ID)
        Next
    End With
    If Me.lvwGroup.ListItems.count = 0 Then
        MsgBox "没有匹配的抗生素！", vbInformation, gstrSysName
        Me.txtFind.SetFocus
    Else
        Me.vfgGroup.SetFocus
    End If
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    mlngItemID = 0: mintGroup = 0
    
    Me.picName.BackColor = Me.BackColor
    Me.picSingle.BackColor = Me.BackColor
    Me.picGroup.BackColor = Me.BackColor
    
    Me.picName.Left = 0: Me.picName.Top = 0
    Me.picSingle.Left = 0: Me.picSingle.Top = Me.picName.Height
    
    With Me.cbo药敏方法
        .AddItem "MIC": .AddItem "DISK": .AddItem "K-B"
    End With
    With Me.lvwGroup.ColumnHeaders
        .Clear
        .Add , "_编码", "编码", 900
        .Add , "_中文名", "中文名", 3000
        .Add , "_缩写", "缩写", 1000
    End With
    With Me.lvwGroup
        .SortKey = .ColumnHeaders("_编码").Index - 1
        .SortOrder = lvwAscending
    End With
    Me.vfgGroup.ZOrder 0

End Sub

Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    Me.picSingle.Height = Me.ScaleHeight - Me.picSingle.Top
    Me.picGroup.Top = Me.ScaleHeight - Me.picGroup.Height - 105
    If Me.Tag <> "" Then
        Me.vfgGroup.Height = Me.picGroup.Top - Me.vfgGroup.Top
    Else
        Me.vfgGroup.Height = Me.ScaleHeight - Me.vfgGroup.Top - 105
    End If
End Sub

Private Sub lvwGroup_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With Me.lvwGroup
        If .SortKey = ColumnHeader.Index - 1 Then
            .SortOrder = IIf(.SortOrder = lvwAscending, lvwDescending, lvwAscending)
        Else
            .SortKey = ColumnHeader.Index - 1
            .SortOrder = lvwAscending
        End If
    End With
End Sub

Private Sub lvwGroup_DblClick()
    Call cmdEdit_Click(0)
End Sub

Private Sub picGroup_Resize()
    Err = 0: On Error Resume Next
    Me.lvwGroup.Height = Me.picGroup.ScaleHeight - Me.lvwGroup.Top
End Sub

Private Sub txtFind_GotFocus()
    Me.txtFind.SelStart = 0: Me.txtFind.SelLength = 1000
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call cmdFind_Click
End Sub

Private Sub txtWHONET码_GotFocus()
    Me.txtWHONET码.SelStart = 0: Me.txtWHONET码.SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txtWHONET码_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(Trim(GCST_INVALIDCHAR), Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt编码_GotFocus()
    Me.txt编码.SelStart = 0: Me.txt编码.SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt编码_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt尿药浓度_GotFocus(Index As Integer)
    Me.txt尿药浓度(Index).SelStart = 0: Me.txt尿药浓度(Index).SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt尿药浓度_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(Trim(GCST_INVALIDCHAR), Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt说明_GotFocus()
    Me.txt说明.SelStart = 0: Me.txt说明.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt说明_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt缩写_GotFocus()
    Me.txt缩写.SelStart = 0: Me.txt缩写.SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt缩写_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(Trim(GCST_INVALIDCHAR), Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt血药浓度_GotFocus(Index As Integer)
    Me.txt血药浓度(Index).SelStart = 0: Me.txt血药浓度(Index).SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt血药浓度_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(Trim(GCST_INVALIDCHAR), Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt英文_GotFocus()
    Me.txt英文.SelStart = 0: Me.txt英文.SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt英文_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(Trim(GCST_INVALIDCHAR), Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt用法用量_GotFocus(Index As Integer)
    Me.txt用法用量(Index).SelStart = 0: Me.txt用法用量(Index).SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt用法用量_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(Trim(GCST_INVALIDCHAR), Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt中文_GotFocus()
    Me.txt中文.SelStart = 0: Me.txt中文.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt中文_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub vfgGroup_DblClick()
    If Me.vfgGroup.MouseRow < Me.vfgGroup.FixedRows Then Exit Sub
    If Me.Tag = "" Then Exit Sub
    Call cmdEdit_Click(1)
End Sub

