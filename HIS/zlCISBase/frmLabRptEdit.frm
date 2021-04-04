VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmLabRptEdit 
   BorderStyle     =   0  'None
   Caption         =   "报告模板编辑"
   ClientHeight    =   5580
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6945
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5580
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picEdit 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   900
      Left            =   0
      ScaleHeight     =   900
      ScaleWidth      =   6945
      TabIndex        =   13
      Top             =   4665
      Width           =   6945
      Begin VB.TextBox txt评语 
         Height          =   300
         Left            =   960
         MaxLength       =   60
         TabIndex        =   15
         Top             =   105
         Width           =   5835
      End
      Begin VB.TextBox txt备注 
         Height          =   300
         Left            =   960
         MaxLength       =   60
         TabIndex        =   17
         Top             =   495
         Width           =   5835
      End
      Begin VB.Label lbl评语 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "报告评语"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   150
         TabIndex        =   14
         Top             =   165
         Width           =   720
      End
      Begin VB.Label lbl备注 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "报告备注"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   150
         TabIndex        =   16
         Top             =   555
         Width           =   720
      End
   End
   Begin VB.PictureBox picName 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   1215
      Left            =   15
      ScaleHeight     =   1215
      ScaleWidth      =   6780
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   15
      Width           =   6780
      Begin VB.TextBox txtItem 
         Height          =   300
         Left            =   1620
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   915
         Width           =   4050
      End
      Begin VB.CommandButton cmdItem 
         Caption         =   "选择"
         Height          =   315
         Left            =   5700
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   915
         Width           =   1100
      End
      Begin VB.TextBox txt说明 
         Height          =   300
         Left            =   2520
         MaxLength       =   60
         TabIndex        =   8
         Top             =   495
         Width           =   4245
      End
      Begin VB.TextBox txt名称 
         Height          =   300
         Left            =   2520
         MaxLength       =   60
         TabIndex        =   4
         Top             =   105
         Width           =   4245
      End
      Begin VB.TextBox txt编码 
         Height          =   300
         Left            =   585
         MaxLength       =   13
         TabIndex        =   2
         Top             =   105
         Width           =   1260
      End
      Begin VB.TextBox txt简码 
         Height          =   300
         Left            =   585
         MaxLength       =   10
         TabIndex        =   6
         Top             =   495
         Width           =   1260
      End
      Begin VB.Label lblItem 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "对应诊疗组合项目"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   135
         TabIndex        =   9
         Top             =   975
         Width           =   1440
      End
      Begin VB.Label lbl说明 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "说明"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   2070
         TabIndex        =   7
         Top             =   555
         Width           =   360
      End
      Begin VB.Label lbl名称 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "名称"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   2070
         TabIndex        =   3
         Top             =   165
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
         Left            =   135
         TabIndex        =   1
         Top             =   165
         Width           =   360
      End
      Begin VB.Label lbl简码 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "简码"
         Height          =   180
         Left            =   135
         TabIndex        =   5
         Top             =   555
         Width           =   360
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vfgEdit 
      Height          =   3210
      Left            =   150
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1305
      Width           =   6645
      _cx             =   11721
      _cy             =   5662
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
   Begin MSComctlLib.ListView lvwItems 
      Height          =   4065
      Left            =   1635
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1245
      Visible         =   0   'False
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   7170
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
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   0
      Top             =   900
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
            Picture         =   "frmLabRptEdit.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmLabRptEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngItemID As Long          '当前显示的项目id
Private mbln微生物 As Boolean

Private Enum mCol
    ID = 0:  中文名: 英文名: 单位: 检验结果: 培养描述
End Enum

Dim objItem As ListItem
Dim lngCount As Long

'--------------------------------------------
'以下为窗体公共方法
'--------------------------------------------
Private Sub RecallReport()
    '功能：重新装载报告项目
    Dim rsTemp As New ADODB.Recordset
    
    
    Err = 0: On Error GoTo ErrHand
    
    gstrSql = "Select B.项目类别" & vbNewLine & _
                "From 检验报告项目 A, 检验项目 B" & vbNewLine & _
                "Where A.报告项目id = B.诊治项目id And B.项目类别 = 2 And 诊疗项目id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Val(Me.txtItem.Tag))
    mbln微生物 = Not rsTemp.EOF
    
    If mbln微生物 = False Then
        gstrSql = "Select I.ID, I.中文名, I.英文名, I.单位, C.检验结果" & vbNewLine & _
                "From 检验报告项目 R, 诊治所见项目 I, (Select 项目id, 检验结果 From 检验模板内容 Where 模板ID = [1]) C" & vbNewLine & _
                "Where R.报告项目ID = I.ID And I.ID = C.项目id(+) And R.诊疗项目id = [2]" & vbNewLine & _
                "Order By r.排列序号"
    Else
        gstrSql = "Select I.ID, I.中文名, I.英文名, '' As 单位, C.检验结果,C.培养描述 " & vbNewLine & _
            "From 检验报告项目 R, 检验细菌 I, (Select 细菌id, 检验结果,培养描述 From 检验模板内容 Where 模板id = [1]) C" & vbNewLine & _
            "Where R.细菌id = I.ID And R.细菌id is not null And I.ID = C.细菌id(+) And R.诊疗项目id = [2]" & vbNewLine & _
            "Order By R.排列序号"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, mlngItemID, Val(Me.txtItem.Tag))
    Me.vfgEdit.Clear
    Set Me.vfgEdit.DataSource = rsTemp: Call setListFormat(True)
    If Me.vfgEdit.Rows > Me.vfgEdit.FixedRows Then Me.vfgEdit.Row = Me.vfgEdit.FixedRows
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub setListFormat(Optional blnKeepData As Boolean)
    '功能：初始化设置列表
    '参数： blnKeepData-是否保留数据，即只是重新设置格式
    With Me.vfgEdit
        .Redraw = flexRDNone
        If blnKeepData = False Then
            .Clear
            .Rows = 1: .FixedRows = 1: .Cols = 6: .FixedCols = 0
        End If
        .TextMatrix(0, mCol.ID) = "ID": .TextMatrix(0, mCol.中文名) = "中文名": .TextMatrix(0, mCol.英文名) = "英文名"
        .TextMatrix(0, mCol.单位) = "单位": .TextMatrix(0, mCol.检验结果) = "检验结果"
'        Call IIf(mbln微生物 = True, .TextMatrix(0, mCol.培养描述) = "培养描述", "")
        
        .ColWidth(mCol.ID) = 0: .ColWidth(mCol.中文名) = 3000: .ColWidth(mCol.英文名) = 1000
        .ColWidth(mCol.单位) = 700: .ColWidth(mCol.检验结果) = 900
'        Call IIf(mbln微生物 = False, .ColWidth(mCol.培养描述) = 0, .ColWidth(mCol.培养描述) = 500)
        For lngCount = 0 To .Cols - 1
            .FixedAlignment(lngCount) = flexAlignCenterCenter
            If .ColWidth(lngCount) = 0 Then .ColHidden(lngCount) = True
        Next
        .Redraw = flexRDDirect
    End With
End Sub

Public Function zlRefresh(lngItemId As Long) As Boolean
    '功能：根据项目id刷新当前显示内容
    Dim rsTemp As New ADODB.Recordset
    
    mlngItemID = lngItemId
    
    '清除此前项目的显示
    Me.txt编码.Text = "": Me.txt名称.Text = "": Me.txt简码.Text = "": Me.txt说明.Text = ""
    Me.txt评语.Text = "": Me.txt备注.Text = ""
    
    '获取指定项目的信息
    Err = 0: On Error GoTo ErrHand
    
    gstrSql = "Select 编码, 名称, 简码, 说明, 诊疗项目id, 检验评语, 检验备注 From 检验模板目录 L Where ID = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemId)
    With rsTemp
        Me.txt编码.MaxLength = .Fields("编码").DefinedSize
        Me.txt名称.MaxLength = .Fields("名称").DefinedSize
        Me.txt简码.MaxLength = .Fields("简码").DefinedSize
        Me.txt说明.MaxLength = .Fields("说明").DefinedSize
        Me.txt评语.MaxLength = .Fields("检验评语").DefinedSize
        Me.txt备注.MaxLength = .Fields("检验备注").DefinedSize
        If .RecordCount > 0 Then
            Me.txt编码.Text = "" & !编码
            Me.txt名称.Text = "" & !名称: Me.txt简码.Text = "" & !简码: Me.txt说明.Text = "" & !说明
            Me.txt评语.Text = "" & !检验评语: Me.txt备注.Text = "" & !检验备注
            For Each objItem In Me.lvwItems.ListItems
                If Mid(objItem.Key, 2) = Val("" & !诊疗项目id) Then
                    objItem.Selected = True
                    Me.txtItem.Tag = Mid(objItem.Key, 2)
                    Me.txtItem.Text = objItem.Text
                    
                End If
            Next
        Else
            Me.txtItem.Tag = ""
            Me.txtItem.Text = ""
        End If
    End With
    Call RecallReport
    
    zlRefresh = True: Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlRefresh = False: Exit Function
End Function

Public Function zlEditStart(blnAdd As Boolean, lngItemId As Long) As Boolean
    '功能：开始项目编辑
    '参数： blnAdd-是否增加，否则为修改
    '       lngItemId-增加的参照项目，或者指定编辑的项目
    Dim rsTemp As New ADODB.Recordset
    
    If blnAdd Then
        Err = 0: On Error GoTo ErrHand
        gstrSql = "Select Nvl(Max(To_Number(编码)), 0) As 编码, Nvl(Max(Length(编码)), 0) As 长度 From 检验模板目录"
'            If .State = adStateOpen Then .Close
'            Call SQLTest(App.ProductName, Me.Caption, gstrSql)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "zlEditStart")
'            Call SQLTest
        With rsTemp
            If !长度 <> 0 And !长度 <= Me.txt编码.MaxLength Then
                Me.txt编码.Text = Format(Val(!编码) + 1, String(!长度, "0"))
            Else
                Me.txt编码.Text = Format(Val(!编码) + 1, String(Me.txt编码.MaxLength, "0"))
            End If
            
            Me.txt名称.Text = "": Me.txt简码.Text = "": Me.txt说明.Text = ""
            Me.txt评语.Text = "": Me.txt备注.Text = ""
            Me.txtItem.Tag = "": Me.txtItem.Text = "": Call setListFormat
        End With
    End If

    Me.Tag = IIf(blnAdd, "增加", "修改")
    Me.BackColor = RGB(250, 250, 250): Me.picName.BackColor = Me.BackColor: Me.picEdit.BackColor = Me.BackColor
    Me.picName.Enabled = True: Me.picEdit.Enabled = True
    Me.vfgEdit.Editable = flexEDKbd: Me.vfgEdit.FocusRect = flexFocusHeavy
    
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
    Me.BackColor = &H8000000F: Me.picName.BackColor = Me.BackColor: Me.picEdit.BackColor = Me.BackColor
    Me.picName.Enabled = False: Me.picEdit.Enabled = False
    Me.vfgEdit.Editable = flexEDNone: Me.vfgEdit.FocusRect = flexFocusNone
    Call Me.zlRefresh(mlngItemID)
End Sub

Public Function zlEditSave() As Long
    '功能：保存正在进行的编辑,并返回正在编辑项目id,保存失败返回0
    Dim lngNewId As Long, strLists As String
    Dim str描述  As String
    '一般特性检查
    If Trim(Me.txt编码.Text) = "" Then
        MsgBox "请输入编码！", vbInformation, gstrSysName
        Me.txt编码.SetFocus: zlEditSave = 0: Exit Function
    End If
    If Val(Me.txt编码.Text) > Val(String(Me.txt编码.MaxLength, "9")) Then
        MsgBox "编码太大！", vbInformation, gstrSysName
        Me.txt编码.SetFocus: zlEditSave = 0: Exit Function
    End If
    If Trim(Me.txt名称.Text) = "" Then
        MsgBox "请输入名称！", vbInformation, gstrSysName
        Me.txt名称.SetFocus: zlEditSave = 0: Exit Function
    End If
    If LenB(StrConv(Trim(Me.txt名称.Text), vbFromUnicode)) > Me.txt名称.MaxLength Then
        MsgBox "名称超长（最多" & Me.txt名称.MaxLength & "个字符或等长汉字）！", vbInformation, gstrSysName
        Me.txt名称.SetFocus: zlEditSave = 0: Exit Function
    End If
    If LenB(StrConv(Trim(Me.txt简码.Text), vbFromUnicode)) > Me.txt简码.MaxLength Then
        MsgBox "缩写超长（最多" & Me.txt简码.MaxLength & "个字符）！", vbInformation, gstrSysName
        Me.txt简码.SetFocus: zlEditSave = 0: Exit Function
    End If
    If LenB(StrConv(Trim(Me.txt说明.Text), vbFromUnicode)) > Me.txt说明.MaxLength Then
        MsgBox "说明超长（最多" & Me.txt说明.MaxLength & "个字符）！", vbInformation, gstrSysName
        Me.txt说明.SetFocus: zlEditSave = 0: Exit Function
    End If
    If LenB(StrConv(Trim(Me.txt评语.Text), vbFromUnicode)) > Me.txt评语.MaxLength Then
        MsgBox "评语超长（最多" & Me.txt评语.MaxLength & "个字符）！", vbInformation, gstrSysName
        Me.txt评语.SetFocus: zlEditSave = 0: Exit Function
    End If
    If LenB(StrConv(Trim(Me.txt备注.Text), vbFromUnicode)) > Me.txt备注.MaxLength Then
        MsgBox "备注超长（最多" & Me.txt备注.MaxLength & "个字符）！", vbInformation, gstrSysName
        Me.txt备注.SetFocus: zlEditSave = 0: Exit Function
    End If
    
    strLists = ""
    With Me.vfgEdit
        For lngCount = .FixedRows To .Rows - 1
            If LenB(StrConv(Trim(.TextMatrix(lngCount, mCol.检验结果)), vbFromUnicode)) > 50 Then
                MsgBox "第" & lngCount & "行结果填写错误！", vbInformation, gstrSysName
                .SetFocus: zlEditSave = 0: Exit Function
            End If
            If mbln微生物 Then
                str描述 = .TextMatrix(lngCount, mCol.培养描述)
            Else
                str描述 = ""
            End If
            strLists = strLists & "|" & .TextMatrix(lngCount, mCol.ID) & ";" & .TextMatrix(lngCount, mCol.检验结果) & ";" & str描述
        Next
    End With
    If strLists = "" Then
        MsgBox "没有设置模板的报告内容！", vbInformation, gstrSysName
        Me.txt说明.SetFocus: zlEditSave = 0: Exit Function
    End If
    strLists = Mid(strLists, 2)
    
    gstrSql = "'" & Trim(Me.txt编码.Text) & "','" & Trim(Me.txt名称.Text) & "','" & Trim(Me.txt简码.Text) & "'"
    gstrSql = gstrSql & ",'" & Trim(Me.txt说明.Text) & "'," & Val(Me.txtItem.Tag)
    gstrSql = gstrSql & ",'" & Trim(Me.txt评语.Text) & "','" & Trim(Me.txt备注.Text) & "'"
    gstrSql = gstrSql & ",'" & strLists & "'"
    
    '数据保存语句组织
    
    lngNewId = mlngItemID
    If Me.Tag = "增加" Then
        lngNewId = zlDatabase.GetNextId("检验模板目录")
        gstrSql = "Zl_检验模板目录_Edit(1," & lngNewId & "," & gstrSql & ")"
    Else
        gstrSql = "Zl_检验模板目录_Edit(2," & lngNewId & "," & gstrSql & ")"
    End If
    
    Err = 0: On Error GoTo ErrHand
    Call SQLTest(App.ProductName, Me.Caption, gstrSql): gcnOracle.Execute gstrSql, , adCmdStoredProc: Call SQLTest
    
    If Me.Tag = "增加" Then mlngItemID = lngNewId
    
    Me.Tag = ""
    Me.BackColor = &H8000000F: Me.picName.BackColor = Me.BackColor: Me.picEdit.BackColor = Me.BackColor
    Me.picName.Enabled = False: Me.picEdit.Enabled = False
    Me.vfgEdit.Editable = flexEDNone: Me.vfgEdit.FocusRect = flexFocusNone
    
    zlEditSave = mlngItemID: Exit Function
    
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlEditSave = 0: Exit Function
End Function

'--------------------------------------------
'以下为窗体控件响应事件
'--------------------------------------------

Private Sub cmdItem_Click()
    Dim rsTemp As New ADODB.Recordset
    With Me.lvwItems
        .Tag = Me.txtItem.Name
        .Left = Me.txtItem.Left
        .Top = Me.txtItem.Top + Me.txtItem.Height
        .ZOrder 0: .Visible = True
        .SetFocus
    End With
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyEscape Then Exit Sub
    If Me.lvwItems.Visible Then
        Me.lvwItems.Visible = False: Me.txtItem.SetFocus: Exit Sub
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    
    mlngItemID = 0
   
    Me.picName.BackColor = Me.BackColor
    Me.picEdit.BackColor = Me.BackColor
    Call setListFormat
    Me.vfgEdit.ZOrder 0

    '------------------------------------------
    With Me.lvwItems.ColumnHeaders
        .Clear
        .Add , "名称", "名称", 3500
        .Add , "编码", "编码", 1000
        .Add , "类型", "类型", 1000
    End With
    With Me.lvwItems
        .ColumnHeaders("编码").Position = 1
        .SortKey = .ColumnHeaders("编码").Index - 1
        .SortOrder = lvwAscending
    End With
    
    Err = 0: On Error GoTo ErrHand
'        gstrSql = "Select I.ID, I.操作类型 As 类型, I.编码, I.名称" & vbNewLine & _
            "From 诊疗项目目录 I, 诊疗检验类型 K" & vbNewLine & _
            "Where I.类别 = 'C' And I.操作类型 = K.名称 And I.组合项目 = 1 And" & vbNewLine & _
            "      (I.撤档时间 Is Null Or To_Char(I.撤档时间, 'yyyy-mm-dd') = '3000-01-01')"
    gstrSql = "Select Distinct I.ID, I.操作类型 As 类型, I.编码, I.名称,decode(N.项目类别,2,2,1) as 项目类别 " & vbNewLine & _
            "From 诊疗项目目录 I, 诊疗检验类型 K, 检验报告项目 M, 检验项目 N, 检验细菌 O" & vbNewLine & _
            "Where I.类别 = 'C' And I.操作类型 = K.名称 And I.ID = M.诊疗项目id And (M.报告项目id = N.诊治项目id Or M.细菌id = O.ID) And" & vbNewLine & _
            "      ((N.项目类别 = 1 And I.组合项目 = 1) Or (N.项目类别 = 2 And I.单独应用 = 1)) And" & vbNewLine & _
            "      (I.撤档时间 Is Null Or To_Char(I.撤档时间, 'yyyy-mm-dd') = '3000-01-01')"
            
'        If .State = adStateOpen Then .Close
'        Call SQLTest(App.Title, Me.Caption, gstrSql)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "Form_Load")
'        Call SQLTest
    Me.lvwItems.ListItems.Clear
    With rsTemp
        Do While Not .EOF
            Set objItem = Me.lvwItems.ListItems.Add(, IIf(Val(!项目类别) = 2, "^", "_") & !ID, !名称)
            objItem.Icon = 1: objItem.SmallIcon = 1
            objItem.SubItems(Me.lvwItems.ColumnHeaders("编码").Index - 1) = !编码
            objItem.SubItems(Me.lvwItems.ColumnHeaders("类型").Index - 1) = !类型
            .MoveNext
        Loop
    End With
    
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog

End Sub

Private Sub Form_Resize()
    Err = 0: 'On Error Resume Next
    Me.picEdit.Top = Me.ScaleHeight - Me.picEdit.Height
    Me.vfgEdit.Height = Me.picEdit.Top - Me.vfgEdit.Top
End Sub

Private Sub lvwItems_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If Me.lvwItems.SortKey = ColumnHeader.Index - 1 Then
        Me.lvwItems.SortOrder = IIf(Me.lvwItems.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        Me.lvwItems.SortKey = ColumnHeader.Index - 1
        Me.lvwItems.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwItems_DblClick()
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    With Me.lvwItems
        Me.txtItem.Tag = Mid(.SelectedItem.Key, 2)
        Me.txtItem.Text = .SelectedItem.Text
        Me.lvwItems.Visible = False
        Call RecallReport
        Call zlCommFun.PressKey(vbKeyTab)
    End With
End Sub

Private Sub lvwItems_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn, vbKeySpace
        If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
        Call lvwItems_DblClick
    End Select
End Sub

Private Sub lvwItems_LostFocus()
    Me.lvwItems.Visible = False
End Sub

Private Sub txtItem_GotFocus()
    Me.txtItem.SelStart = 0: Me.txtItem.SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txtItem_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub txt备注_GotFocus()
    Me.txt编码.SelStart = 0: Me.txt编码.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt备注_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
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

Private Sub txt简码_GotFocus()
    Me.txt简码.SelStart = 0: Me.txt简码.SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt简码_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt名称_GotFocus()
    Me.txt名称.SelStart = 0: Me.txt名称.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt名称_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Me.txt名称.Text = MoveSpecialChar(Me.txt名称.Text)
        Me.txt简码.Text = zlStr.GetCodeByORCL(Me.txt名称.Text, False, Me.txt简码.MaxLength)
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    End If
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt评语_GotFocus()
    Me.txt评语.SelStart = 0: Me.txt评语.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt评语_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt说明_GotFocus()
    Me.txt说明.SelStart = 0: Me.txt说明.SelLength = 1000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt说明_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub vfgEdit_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> mCol.检验结果 And Col <> mCol.培养描述 Then Cancel = True
End Sub

