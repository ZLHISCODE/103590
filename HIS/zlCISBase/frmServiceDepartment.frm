VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmServiceDepartment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "科室选择"
   ClientHeight    =   4380
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6075
   Icon            =   "frmServiceDepartment.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   6075
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton cmd取消 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4800
      TabIndex        =   7
      Top             =   3960
      Width           =   1100
   End
   Begin VB.PictureBox picDrug 
      Height          =   2940
      Left            =   1320
      ScaleHeight     =   2880
      ScaleWidth      =   4035
      TabIndex        =   2
      Top             =   960
      Visible         =   0   'False
      Width           =   4095
      Begin VB.CommandButton cmdCancel 
         Caption         =   "取消"
         Height          =   350
         Left            =   1200
         TabIndex        =   6
         Top             =   2280
         Width           =   1100
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "确定"
         Height          =   350
         Left            =   120
         TabIndex        =   5
         Top             =   2280
         Width           =   1100
      End
      Begin VB.CheckBox chk科室 
         Appearance      =   0  'Flat
         Caption         =   "全选"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   3000
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   0
         Width           =   675
      End
      Begin MSComctlLib.ListView lvwItems 
         Height          =   2055
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Visible         =   0   'False
         Width           =   3570
         _ExtentX        =   6297
         _ExtentY        =   3625
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "imgList"
         SmallIcons      =   "imgList"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfDepartment 
      Height          =   3735
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   5655
      _cx             =   9975
      _cy             =   6588
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
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmServiceDepartment.frx":6852
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
   Begin VB.CommandButton cmd确定 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3600
      TabIndex        =   0
      Top             =   3960
      Width           =   1100
   End
End
Attribute VB_Name = "frmServiceDepartment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mstr存储库房 As String
Private mstr存储库房ID As String
Private mstr服务科室 As String
Private mstr服务科室ID As String
Private mstr库房科室 As String
Private mstr库房科室ID As String
Private mstrArr存储库房() As String
Private mstrArr存储库房ID() As String
Private mstrArr服务科室() As String
Private mstrArr库房科室ID() As String
Private mrs科室 As ADODB.Recordset
Private mstr服务对象 As String

Private Enum mSpecColumn
    存储库房 = 0
    存储库房ID = 1
    服务科室 = 2
    服务科室id = 3
End Enum

Public Sub ShowMe(ByVal frmParent As Object, ByVal str存储库房 As String, ByVal str存储库房ID As String, ByVal str库房科室 As String, ByVal str库房科室ID As String)
    mstr存储库房 = str存储库房
    mstr存储库房ID = str存储库房ID
    mstr库房科室 = str库房科室
    mstr库房科室ID = str库房科室ID

    Me.Show 1, frmParent
End Sub

Private Sub cmd取消_Click()
    Unload Me
End Sub

Private Sub cmd确定_Click()
    Dim i As Integer
    
    mstr库房科室 = ""
    mstr库房科室ID = ""
    
    For i = 1 To vsfDepartment.Rows - 1
        If vsfDepartment.TextMatrix(i, 2) = "" Then
            mstr库房科室ID = mstr库房科室ID & "!!" & vsfDepartment.TextMatrix(i, 1) & "|"
        End If
        
        If vsfDepartment.TextMatrix(i, 2) <> "" Then
            mstr库房科室 = mstr库房科室 & "；" & vsfDepartment.TextMatrix(i, 0) & "：" & vsfDepartment.TextMatrix(i, 2)
            mstr库房科室ID = mstr库房科室ID & "!!" & vsfDepartment.TextMatrix(i, 1) & "|" & vsfDepartment.TextMatrix(i, 3)
        End If
    Next
     
     mstr库房科室 = Mid(mstr库房科室, 2)
     mstr库房科室ID = Mid(mstr库房科室ID, 3)
     
    Call frmBatchUpdate.ShowDepartment(mstr库房科室, mstr库房科室ID, 0)
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    chk科室.Value = 0
    picDrug.Visible = False
    vsfDepartment.Enabled = True
End Sub

Private Sub Form_Load()
    Call Init初始化表格
    Call Init初始化库房数据
End Sub

Private Sub Init初始化表格()
    
    VsfGridColFormat vsfDepartment, mSpecColumn.存储库房, "存储库房", 1500, flexAlignLeftCenter, "存储库房"
    VsfGridColFormat vsfDepartment, mSpecColumn.存储库房ID, "存储库房ID", 1500, flexAlignCenterCenter, "存储库房ID"
    VsfGridColFormat vsfDepartment, mSpecColumn.服务科室, "服务科室", 4000, flexAlignLeftCenter, "服务科室"
    VsfGridColFormat vsfDepartment, mSpecColumn.服务科室id, "服务科室id", 4000, flexAlignCenterCenter, "服务科室id"
    vsfDepartment.ColComboList(mSpecColumn.服务科室) = "..."
    
End Sub

Private Sub Init初始化库房数据()
    '功能一：如果批量规格里选择了具体的存储库房后，初始化vsfDepartment里的库房和库房ID(用于查询服务科室)
    '功能二：如果是选择了服务科室之后，再次点击，就会把对应的服务科室也显示出来
    Dim i As Integer, j As Integer
    Dim rsRoom As New ADODB.Recordset
    
    mstrArr存储库房 = Split(mstr存储库房, "|")
    mstrArr存储库房ID = Split(mstr存储库房ID, "!!")
    mstrArr库房科室ID = Split(mstr库房科室ID, "!!")
    
    For i = LBound(mstrArr存储库房) To UBound(mstrArr存储库房)
        vsfDepartment.Rows = vsfDepartment.Rows + 1
        vsfDepartment.RowHeight(i + 1) = 400
        vsfDepartment.TextMatrix(i + 1, mSpecColumn.存储库房) = mstrArr存储库房(i)
    Next
     
    For i = LBound(mstrArr存储库房ID) To UBound(mstrArr存储库房ID)
        vsfDepartment.TextMatrix(i + 1, mSpecColumn.存储库房ID) = Split(mstrArr存储库房ID(i), "|")(0)
    Next
    
    For i = LBound(mstrArr库房科室ID) To UBound(mstrArr库房科室ID)
        For j = 1 To vsfDepartment.Rows - 1
            If Split(mstrArr库房科室ID(i), "|")(0) = vsfDepartment.TextMatrix(j, 1) Then
                    vsfDepartment.TextMatrix(j, 3) = Split(mstrArr库房科室ID(i), "|")(1)
                    
                    gstrSql = "select a.名称 from 部门表 a where a.id in(Select Column_Value From Table(f_num2list([1])))"
                    Set rsRoom = zlDatabase.OpenSQLRecord(gstrSql, "", vsfDepartment.TextMatrix(j, 3))
                    
                    Do While Not rsRoom.EOF
                        vsfDepartment.TextMatrix(j, 2) = vsfDepartment.TextMatrix(j, 2) & "," & rsRoom!名称
                        rsRoom.MoveNext
                    Loop
                    If vsfDepartment.TextMatrix(j, 2) <> "" Then
                        vsfDepartment.TextMatrix(j, 2) = Mid(vsfDepartment.TextMatrix(j, 2), 2)
                    End If
                    
                Exit For
            End If
        Next
    Next
End Sub
Public Sub VsfGridColFormat(ByVal objGrid As VSFlexGrid, ByVal intCol As Integer, ByVal strColName As String, _
    ByVal lngColWidth As Long, ByVal intColAlignment As Integer, _
    Optional ByVal strColKey As String = "", Optional ByVal intFixedColAlignment As Integer = 4)
    'vsf列设置：列名，列宽，列对齐方式，固定列对齐方式（默认为居中对齐）

    With objGrid
        .TextMatrix(0, intCol) = strColName
        .ColWidth(intCol) = lngColWidth
        .ColAlignment(intCol) = intColAlignment
        .ColKey(intCol) = strColKey
        .FixedAlignment(intCol) = intFixedColAlignment
    End With
End Sub


Private Sub vsfDepartment_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Select Case Col
        Case mSpecColumn.服务科室
            If Check服务科室 = False Then
'                Call Init加载科室数据
            Call frmServiceSelect.ShowMe(frmServiceDepartment, vsfDepartment.Row, mstr服务对象, 2)
            End If
    End Select
End Sub

Private Function Check服务科室() As Boolean
    '功能：检查当前库房是不是药房或者是否设置临床科室
    '返回值 true 当前库房不是药房也没有设置临床科室,false 当前库房是药房或者或者设置了临床科室
    Dim str服务对象 As String
    Dim rsTemp As ADODB.Recordset

    On Error GoTo ErrHandle
    str服务对象 = ""
    gstrSql = "select distinct 服务对象 from 部门性质说明 where 部门ID=[1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "提取服务对象", vsfDepartment.TextMatrix(vsfDepartment.Row, 1))

    Do While Not rsTemp.EOF
        str服务对象 = str服务对象 & "," & rsTemp!服务对象
        rsTemp.MoveNext
    Loop
    If str服务对象 <> "" Then
        str服务对象 = Mid(str服务对象, 2)
        If InStr(1, str服务对象, 3) <> 0 Then
            str服务对象 = "0,1,2,3"
        ElseIf InStr(1, str服务对象, 1) <> 0 Or InStr(1, str服务对象, 2) <> 0 Then
            str服务对象 = str服务对象 & ",3"
        End If
    Else
        str服务对象 = "0"
    End If
    mstr服务对象 = str服务对象
    
    gstrSql = "Select distinct a.Id, a.编码, a.名称, a.简码" & vbNewLine & _
            "From 部门表 a, 部门性质说明 b, 部门性质分类 c" & vbNewLine & _
            "Where a.Id = b.部门id And b.工作性质 = c.名称 And Instr('3ABCDEF', c.编码) > 0 And" & vbNewLine & _
            "  (a.撤档时间 Is Null Or a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD')) And Instr([1], ',' || b.服务对象 || ',') > 0 order by id"

    Set mrs科室 = zlDatabase.OpenSQLRecord(gstrSql, "提取服务科室", "," & str服务对象 & ",")

    If mrs科室.RecordCount = 0 Then
        MsgBox "当前库房不是药房或者未设置临床科室！[部门管理]", vbInformation, gstrSysName
        vsfDepartment.Text = ""
        vsfDepartment.TextMatrix(vsfDepartment.Row, vsfDepartment.Col) = ""
        Check服务科室 = True
        Exit Function
    End If
    Check服务科室 = False
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Init加载科室数据()
    '在lvwItems显示出具体的服务科室
    Dim str服务对象 As String
    Dim objItem As ListItem
    Dim intItem As Integer
    Dim i As Integer, j As Integer
    
    Call AddColumnHeader(False)
    Me.lvwItems.ListItems.Clear
    Me.lvwItems.Checkboxes = True
    
    Do While Not mrs科室.EOF
        Set objItem = Me.lvwItems.ListItems.Add(, "_" & mrs科室!ID, mrs科室!名称, , 3)
        objItem.SubItems(Me.lvwItems.ColumnHeaders("编码").Index - 1) = mrs科室!编码
        objItem.SubItems(Me.lvwItems.ColumnHeaders("简码").Index - 1) = mrs科室!简码
        mrs科室.MoveNext
    Loop
    
    With Me.picDrug
        .Left = Me.vsfDepartment.Left + 1500
        .Top = Me.vsfDepartment.Top + Me.vsfDepartment.CellTop
        If .Top + .Height > Me.ScaleHeight Then
            .Top = Me.ScaleHeight - .Height
        End If

        lvwItems.Move 0, 250, picDrug.Width, picDrug.Height - 670
        chk科室.Move 3300, 0
        cmdOk.Move 0, picDrug.Height - 400
        cmdCancel.Move cmdOk.Width, cmdOk.Top
        
        lvwItems.Visible = True
        .ZOrder 0: .Visible = True
        vsfDepartment.Enabled = False
        .SetFocus
    End With
    
    '当选择了服务科室后，再次点击服务科室，会把现有的服务科室在lvwItems显示出来
    mstr服务科室 = vsfDepartment.TextMatrix(vsfDepartment.Row, mSpecColumn.服务科室)
    mstrArr服务科室 = Split(mstr服务科室, ",")
    
    For i = LBound(mstrArr服务科室) To UBound(mstrArr服务科室)
        For intItem = 1 To lvwItems.ListItems.Count
            If mstrArr服务科室(i) = lvwItems.ListItems(intItem).Text Then
                lvwItems.ListItems(intItem).Checked = True
                j = j + 1
            End If
        Next
    Next
    
    If j = lvwItems.ListItems.Count Then
        chk科室.Value = 1
    ElseIf j > 0 And j < lvwItems.ListItems.Count Then
        chk科室.Value = 2
    End If
End Sub

Private Sub AddColumnHeader(Optional ByVal bln药品 As Boolean = True)
 
        With Me.lvwItems.ColumnHeaders
            .Clear
            .Add , "名称", "名称", 2000
            .Add , "编码", "编码", 800
            .Add , "简码", "简码", 700
        End With
        
        With Me.lvwItems
            .Checkboxes = True
            .ColumnHeaders("编码").Position = 1
            .Sorted = False '关闭排序功能
        End With
    
End Sub

Private Sub cmdOK_Click()
    Dim intItem As Integer, intItems As Integer
    
    '把选择的服务科室和服务科室ID显示到vsfDepartment
    mstr服务科室ID = ""
    mstr服务科室 = ""
    intItems = Me.lvwItems.ListItems.Count
    For intItem = 1 To intItems
        If lvwItems.ListItems(intItem).Checked Then
            mstr服务科室ID = mstr服务科室ID & "," & Mid(lvwItems.ListItems(intItem).Key, 2)
            mstr服务科室 = mstr服务科室 & "," & Mid(lvwItems.ListItems(intItem).Text, 1)
        End If
    Next
    
    mstr服务科室 = Mid(mstr服务科室, 2)
    mstr服务科室ID = Mid(mstr服务科室ID, 2)
    If vsfDepartment.Row <> 0 Then
        vsfDepartment.TextMatrix(vsfDepartment.Row, mSpecColumn.服务科室) = mstr服务科室
        vsfDepartment.TextMatrix(vsfDepartment.Row, mSpecColumn.服务科室id) = mstr服务科室ID
    End If
    
    picDrug.Visible = False
    vsfDepartment.Enabled = True
End Sub

Private Sub vsfDepartment_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If Col = mSpecColumn.服务科室 Then
        KeyAscii = 0
    End If
End Sub
Private Sub vsfDepartment_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = mSpecColumn.存储库房 Then
        Cancel = True
    End If
End Sub

Private Sub chk科室_Click()
'库房全选按钮
    If chk科室.Value = 2 Then Exit Sub
    Call SetSelect(lvwItems, chk科室.Value)
End Sub
Private Sub SetSelect(ByVal lvwObj As Object, Optional ByVal BlnSelect As Boolean = True)
'全选功能
    Dim intSelect As Integer
    With lvwObj
        For intSelect = 1 To .ListItems.Count
            .ListItems(intSelect).Checked = BlnSelect
        Next
    End With
End Sub
Private Sub lvwItems_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'具体选择的存储库房
    Call ItemCheck(lvwItems, Item, chk科室)
End Sub
Private Sub ItemCheck(ByVal lvwObj As Object, ByVal Item As MSComctlLib.ListItem, ByVal chkObj As CheckBox)
'纪录选择的库房
    Dim lngCheck As Long, blnCheck As Boolean, intCount As Integer
    
    intCount = 0
    With lvwObj
        For lngCheck = 1 To .ListItems.Count
            If .ListItems(lngCheck).Checked = True Then
                intCount = intCount + 1
            End If
        Next
        
        If intCount = lvwObj.ListItems.Count Then
            chkObj.Value = 1
        ElseIf intCount > 0 Then
            chkObj.Value = 2
        Else
            chkObj.Value = 0
        End If
    End With
End Sub
