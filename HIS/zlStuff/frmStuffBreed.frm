VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStuffBreed 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "卫生材料品种编辑"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7980
   Icon            =   "frmStuffBreed.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   7980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton cmdSaveAddSpec 
      Caption         =   "保存后新增规格(&B)"
      Height          =   350
      Left            =   3360
      TabIndex        =   29
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CommandButton cmdSaveAddItem 
      Caption         =   "保存后新增品种(&A)"
      Height          =   350
      Left            =   1560
      TabIndex        =   28
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CommandButton cmd分类 
      Caption         =   "…"
      Height          =   285
      Left            =   7440
      TabIndex        =   27
      TabStop         =   0   'False
      Tag             =   "分类"
      ToolTipText     =   "按*打开选择器"
      Top             =   728
      Width           =   285
   End
   Begin VB.ComboBox cbo适用性别 
      Height          =   300
      Left            =   5475
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Top             =   2280
      Width           =   2220
   End
   Begin VB.ComboBox cmbStationNo 
      Height          =   300
      Left            =   1515
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   4200
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VSFlex8Ctl.VSFlexGrid vsEditBill 
      Height          =   1410
      Left            =   1515
      TabIndex        =   19
      Top             =   2715
      Width           =   6165
      _cx             =   10874
      _cy             =   2487
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
      BackColorBkg    =   -2147483634
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmStuffBreed.frx":030A
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
   Begin VB.ComboBox cbo单位 
      Height          =   300
      Left            =   1515
      TabIndex        =   11
      Top             =   2295
      Width           =   2205
   End
   Begin VB.Frame fra 
      Height          =   60
      Left            =   0
      TabIndex        =   18
      Top             =   600
      Width           =   8115
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6555
      TabIndex        =   24
      Top             =   4800
      Width           =   1100
   End
   Begin VB.CommandButton cmd帮助 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   255
      Picture         =   "frmStuffBreed.frx":03A0
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   4800
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "保存退出(&O)"
      Height          =   350
      Left            =   5160
      TabIndex        =   23
      Top             =   4800
      Width           =   1215
   End
   Begin VB.TextBox txt英文 
      Height          =   300
      Left            =   5520
      MaxLength       =   40
      TabIndex        =   13
      Top             =   1125
      Width           =   2175
   End
   Begin VB.TextBox txt五笔 
      Height          =   300
      Left            =   4800
      MaxLength       =   12
      TabIndex        =   9
      Top             =   1920
      Width           =   2340
   End
   Begin VB.TextBox txt拼音 
      Height          =   300
      Left            =   1515
      MaxLength       =   12
      TabIndex        =   8
      Top             =   1920
      Width           =   2160
   End
   Begin VB.TextBox txt名称 
      Height          =   300
      Left            =   1515
      MaxLength       =   40
      TabIndex        =   6
      Top             =   1515
      Width           =   6175
   End
   Begin VB.TextBox txt编码 
      Height          =   300
      Left            =   1515
      MaxLength       =   13
      TabIndex        =   4
      Top             =   1125
      Width           =   2175
   End
   Begin VB.Frame fraLine 
      Height          =   30
      Left            =   -15
      TabIndex        =   16
      Top             =   4620
      Width           =   8490
   End
   Begin MSComctlLib.TreeView tvwClass 
      Height          =   3570
      Left            =   450
      TabIndex        =   17
      TabStop         =   0   'False
      Tag             =   "1000"
      Top             =   6375
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   6297
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "imgList"
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   4395
      Top             =   6375
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
            Picture         =   "frmStuffBreed.frx":04EA
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffBreed.frx":0A84
            Key             =   "expend"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffBreed.frx":101E
            Key             =   "ItemUse"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffBreed.frx":15B8
            Key             =   "ItemStop"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txt分类 
      Height          =   300
      Left            =   1515
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   2
      Top             =   720
      Width           =   5895
   End
   Begin VB.Label lbl适用性别 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "适用性别(&X)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   4440
      TabIndex        =   26
      Top             =   2355
      Width           =   990
   End
   Begin VB.Label lblStationNo 
      AutoSize        =   -1  'True
      Caption         =   "院区"
      Height          =   180
      Left            =   795
      TabIndex        =   20
      Top             =   4275
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lbl别名 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "其他别名(&Q)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   435
      TabIndex        =   22
      Top             =   2745
      Width           =   990
   End
   Begin VB.Label Lbl单位 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "散装单位(&U)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   435
      TabIndex        =   10
      Top             =   2355
      Width           =   990
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "注：该品种建立于2003-09-01"
      Height          =   180
      Left            =   3915
      TabIndex        =   14
      Top             =   4305
      Width           =   2580
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "设置卫生材料的相关品种."
      Height          =   180
      Left            =   825
      TabIndex        =   0
      Top             =   240
      Width           =   2070
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   0
      Left            =   165
      Picture         =   "frmStuffBreed.frx":1B52
      Top             =   60
      Width           =   480
   End
   Begin VB.Label lbl英文 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "英文名称(&E)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   4440
      TabIndex        =   12
      Top             =   1185
      Width           =   990
   End
   Begin VB.Label lbl分类 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "材料分类(&F)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   435
      TabIndex        =   1
      Top             =   825
      Width           =   990
   End
   Begin VB.Label lbl简码 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "名称简码(&S)                         (拼音)                                (五笔)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   480
      TabIndex        =   7
      Top             =   1980
      Width           =   7200
   End
   Begin VB.Label lbl名称 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "通用名称(&N)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   435
      TabIndex        =   5
      Top             =   1575
      Width           =   990
   End
   Begin VB.Label lbl编码 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "编码(&D)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   795
      TabIndex        =   3
      Top             =   1185
      Width           =   630
   End
End
Attribute VB_Name = "frmStuffBreed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mstr诊疗ID As String         '当前编辑的材料ID
Dim mlng分类id As Long

Dim mintSuccess As Integer
Dim mintEditType As gEditType    '编辑类型
Dim mblnChange As Boolean
Dim mstrPrivs As String         '权限串
Dim mblnFrist As Boolean        '第一次运行系统时
Dim mintCount As Integer
Dim mintCodeLength As Integer   '编码的长度,从数据库中读取出来的长度
Private mlng品种id As Long      '记录品种id
Private Const mlngModule = 1711


Private Sub GetDefineSize()
    '------------------------------------------------------------------------------------------------------------------
    '功能：得到数据库的表字段的长度
    '编制:刘兴宏
    '日期:2007/05/24
    '------------------------------------------------------------------------------------------------------------------
    On Error GoTo ErrHandle
    Dim rsTmp As New ADODB.Recordset
    gstrSQL = "Select 编码,名称 From 诊疗项目目录 Where Rownum<0"
    Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, Me.Caption)
    mintCodeLength = rsTmp.Fields("编码").DefinedSize
    txt编码.MaxLength = rsTmp.Fields("编码").DefinedSize
    txt名称.MaxLength = rsTmp.Fields("名称").DefinedSize
    
    gstrSQL = "Select 简码,名称 From 诊疗项目别名 Where Rownum<0"
    Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, Me.Caption)
    txt拼音.MaxLength = rsTmp.Fields("简码").DefinedSize
    txt五笔.MaxLength = txt拼音.MaxLength
    txt英文.MaxLength = rsTmp.Fields("名称").DefinedSize
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Function ShowEditCard(ByVal frmMain As Object, _
    intEditType As gEditType, Optional ByVal str诊疗id As String = "", Optional ByVal lng分类id As Long, Optional strPrivs As String = "") As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:编辑卫生材料
    '--入参数:frmMain-调用的主窗体
    '--       intEditType -编辑类型
    '--       str诊疗ID-编辑档案的当前诊疗ID
    '         strPrivs-权限串
    '--出参数:
    '--返  回:编辑成功,返回ture,否则false
    '编制:刘兴宏
    '日期;2007/05/24
    '-----------------------------------------------------------------------------------------------------------
    
    Dim rsTemp As New ADODB.Recordset
    Dim intTemp As Byte
    Dim strTemp As String
    
    mlng分类id = lng分类id
    mstr诊疗ID = str诊疗id
    mstrPrivs = strPrivs
    mintEditType = intEditType
    mintSuccess = 0
    
    frmStuffBreed.Show 1, frmMain
    ShowEditCard = mintSuccess > 0
End Function

Private Sub cbo单位_Change()
        mblnChange = True
End Sub

Private Sub cbo单位_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If

End Sub

Private Sub cmdSaveAddItem_Click()
    Call cmdOK_Click
End Sub

Private Sub cmdSaveAddSpec_Click()
    Call cmdOK_Click
End Sub

Private Sub cmd分类_Click()
    If Me.tvwClass.Nodes.Count = 0 Then
        Call Load诊疗分类信息
    End If
   With Me.tvwClass
        .Left = Me.txt分类.Left
        .Top = Me.txt分类.Top + Me.txt分类.Height
        .Width = txt分类.Width
        .ZOrder 0: .Visible = True
        .SetFocus
    End With
End Sub
Private Sub cbo单位_LostFocus()
    Dim strTmp As String
    Dim i As Long
    Dim blnAdd As Boolean
    ImeLanguage False
    
    strTmp = cbo单位.Text
    blnAdd = True
    For i = 0 To cbo单位.ListCount - 1
        If cbo单位.List(i) = Trim(strTmp) Then
            blnAdd = False
            Exit For
        End If
    Next
    If blnAdd And strTmp <> "" Then
        cbo单位.AddItem strTmp
    End If
    
End Sub

Private Sub cbo单位_GotFocus()
    Me.cbo单位.SelStart = 0: Me.cbo单位.SelLength = 100
    ImeLanguage True
End Sub

Private Sub cbo单位_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
       Exit Sub
    Case Else
        zlControl.TxtCheckKeyPress cbo单位, KeyAscii, m文本式
    End Select
End Sub


Private Sub InitCardData(ByVal lng诊疗ID As Long)
    '------------------------------------------------------------------------------------------------------------------
    '功能：初始化卫生材料品种的卡片数据
    '参数:lng诊疗-指定的诊疗ID
    '编制:刘兴宏
    '日期:2007/05/24
    '------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim i As Long
    
    On Error GoTo ErrHandle
    Me.lblNote.Caption = ""
    If mintEditType <> g查看 Then
        gstrSQL = "select distinct 计算单位 from 诊疗项目目录 where 类别 ='4' and 计算单位 is not null"
        zlDatabase.OpenRecordset rsTemp, gstrSQL, "获取计算单位"
        With rsTemp
            cbo单位.Clear
            Do While Not .EOF
                Me.cbo单位.AddItem .Fields(0).Value
                .MoveNext
            Loop
        End With
    End If
    
    Me.cbo适用性别.Clear
    Me.cbo适用性别.AddItem "0-无性别区分"
    Me.cbo适用性别.AddItem "1-男性"
    Me.cbo适用性别.AddItem "2-女性"
    Me.cbo适用性别.ListIndex = 0
    
    Me.vsEditBill.Clear 1
    Me.vsEditBill.Rows = 2
    If mintEditType = g新增 Then
        Me.tvwClass.Nodes("_" & mlng分类id).Selected = True
        Me.txt分类.Text = Me.tvwClass.SelectedItem.Text
        Me.txt分类.Tag = mlng分类id
        Me.txt编码.Text = GetMaxCode()
        Me.txt名称.Text = ""
        Me.txt拼音.Text = ""
        Me.txt五笔.Text = ""
        Me.txt英文.Text = ""
        Me.cbo单位.Text = ""
        Exit Sub
    End If

    '基本信息项目
    gstrSQL = "select I.分类ID,I.编码,I.名称,I.计算单位," & _
            "        I.建档时间,nvl(I.撤档时间,to_date('3000-01-01','YYYY-MM-DD')) as 撤档时间,I.站点,Nvl(I.适用性别,0) As 适用性别 " & _
            " from 诊疗项目目录 I" & _
            " where  I.ID=[1]   "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng诊疗ID)
        
    With rsTemp
        If Not .EOF Then
            With cmbStationNo
                For i = 1 To .ListCount - 1
                    If Mid(cmbStationNo.List(i), 1, InStr(1, cmbStationNo.List(i), "-") - 1) = zlStr.nvl(rsTemp!站点) Then
                        .ListIndex = i: Exit For
                    End If
                Next
            End With
            Me.lblNote.Caption = "注：该材料建立于" & Format(!建档时间, "YYYY-MM-DD")
            If Format(!撤档时间, "YYYY-MM-DD") <> "3000-01-01" Then
                Me.lblNote.Caption = Me.lblNote.Caption & "，于" & Format(!撤档时间, "YYYY-MM-DD") & "停用。"
            End If
            
            Me.tvwClass.Nodes("_" & !分类id).Selected = True
            Me.txt分类.Text = Me.tvwClass.SelectedItem.Text
            Me.txt分类.Tag = Mid(Me.tvwClass.SelectedItem.Key, 2)
            mlng分类id = Val(Me.txt分类.Tag)
            Me.txt编码.Text = !编码
            Me.txt名称.Text = !名称
            Me.cbo单位.Text = zlStr.nvl(!计算单位)
            Me.cbo适用性别.ListIndex = !适用性别
        End If
    End With
       
    '正名简码与英文名
    gstrSQL = "select 名称,性质,简码,码类 from 诊疗项目别名 where 性质 in (1,2) and 诊疗项目ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng诊疗ID)
    With rsTemp
        Do While Not .EOF
            If !性质 = 1 And !码类 = 1 Then Me.txt拼音.Text = zlStr.nvl(!简码)
            If !性质 = 1 And !码类 = 2 Then Me.txt五笔.Text = zlStr.nvl(!简码)
            If !性质 = 2 Then Me.txt英文.Text = zlStr.nvl(!名称)
            .MoveNext
        Loop
    End With
    
    '其他别名
    gstrSQL = "select N.名称,P.简码 as 拼音,W.简码 as 五笔" & _
            " from (select distinct 名称 from 诊疗项目别名 where 诊疗项目ID=[1] and 性质=9) N," & _
            "      (select 名称,简码 from 诊疗项目别名 where 诊疗项目ID=[1] and 性质=9 and 码类=1) P," & _
            "      (select 名称,简码 from 诊疗项目别名 where 诊疗项目ID=[1] and 性质=9 and 码类=2) W" & _
            " where N.名称=P.名称(+) and N.名称=W.名称(+)"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng诊疗ID)
    
    With rsTemp
        Do While Not .EOF
            If Me.vsEditBill.Rows - 1 < .AbsolutePosition Then Me.vsEditBill.Rows = Me.vsEditBill.Rows + 1
            Me.vsEditBill.TextMatrix(.AbsolutePosition, 0) = .AbsolutePosition
            Me.vsEditBill.TextMatrix(.AbsolutePosition, 1) = zlStr.nvl(!名称)
            Me.vsEditBill.TextMatrix(.AbsolutePosition, 2) = zlStr.nvl(!拼音)
            Me.vsEditBill.TextMatrix(.AbsolutePosition, 3) = zlStr.nvl(!五笔)
            .MoveNext
        Loop
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetEditCtrEnable()
    '------------------------------------------------------------------------------------------------------------------
    '功能：设置编辑控件的Enable属性
    '编制:刘兴宏
    '日期:2007/05/24
    '------------------------------------------------------------------------------------------------------------------
    
    Dim blnStuffModify As Boolean
    
    If mintEditType = g新增 Or mintEditType = g修改 Then
        blnStuffModify = True
    Else
        cmdOK.Visible = False
    End If
    Me.txt分类.Enabled = blnStuffModify
    Me.txt编码.Enabled = blnStuffModify
    Me.txt名称.Enabled = blnStuffModify
    Me.cmd分类.Enabled = blnStuffModify
    Me.txt拼音.Enabled = blnStuffModify
    Me.txt五笔.Enabled = blnStuffModify
    Me.txt英文.Enabled = blnStuffModify
    Me.cbo单位.Enabled = blnStuffModify
    Me.cmbStationNo.Enabled = blnStuffModify
    Me.cbo适用性别.Enabled = blnStuffModify
    If blnStuffModify Then
        vsEditBill.Editable = flexEDKbdMouse
    Else
        vsEditBill.Editable = flexEDNone
    End If
    
    If blnStuffModify = False Then
        SetCtlBackColor txt分类
        SetCtlBackColor txt编码
        SetCtlBackColor txt名称
        SetCtlBackColor txt拼音
        SetCtlBackColor txt英文
        SetCtlBackColor txt五笔
    End If
End Sub


Private Sub Form_Activate()

    If mblnFrist = False Then Exit Sub
    mblnFrist = False
    
    '初始站点
    cmbStationNo.Visible = gSystem_Para.bln存在站点
    lblStationNo.Visible = cmbStationNo.Visible
    
    
    '----------设置相关的输入长度-------------------------------------
    Call GetDefineSize
     
    '取诊疗分类目录数据
    Call Load诊疗分类信息
     
    '----------初始卡片数据-------------------------------------
    Call InitCardData(Val(mstr诊疗ID))
    
    '设置编辑控件
    Call SetEditCtrEnable
    If txt名称.Enabled Then txt名称.SetFocus
End Sub

Private Function ISValied() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:验证数据的合法性
    '--入参数:
    '--出参数:
    '--返  回:合法,返回true,否则返回False
    '--编制:刘兴宏
    '--日期:2007/05/24
    '-----------------------------------------------------------------------------------------------------------
    Dim i As Integer
    Dim strTmp As String, strTemp As String
    Dim strName As String
    
    ISValied = False
  '编辑数据检查
    If Trim(Me.txt编码.Text) = "" Then
        MsgBox "请输入材料编码！", vbInformation, gstrSysName
        Me.txt编码.SetFocus: Exit Function
    End If
    If LenB(StrConv(Trim(Me.txt编码.Text), vbFromUnicode)) > mintCodeLength Then
        MsgBox "编码的长度超长（最多" & mintCodeLength & "个字符）！", vbInformation, gstrSysName
        Me.txt编码.SetFocus: Exit Function
    End If
    If Trim(Me.txt名称.Text) = "" Then
        MsgBox "请输入材料名称！", vbInformation, gstrSysName
        Me.txt名称.SetFocus: Exit Function
    End If
    
    If LenB(StrConv(Trim(Me.txt名称.Text), vbFromUnicode)) > txt名称.MaxLength Then
        MsgBox "材料名称长度超长（最多" & txt名称.MaxLength & "个字符或" & txt名称.MaxLength \ 2 & "个汉字）！", vbInformation, gstrSysName
        Me.txt名称.SetFocus: Exit Function
    End If
    
    If LenB(StrConv(Trim(Me.txt拼音.Text), vbFromUnicode)) > txt拼音.MaxLength Then
        MsgBox "材料拼音简码长度超长（最多" & txt拼音.MaxLength & "个字符或" & txt拼音.MaxLength \ 2 & "个汉字）！", vbInformation, gstrSysName
        Me.txt拼音.SetFocus: Exit Function
    End If
    If LenB(StrConv(Trim(Me.txt五笔.Text), vbFromUnicode)) > txt五笔.MaxLength Then
        MsgBox "材料五笔简码长度超长（最多" & txt五笔.MaxLength & "个字符或" & txt五笔.MaxLength \ 2 & "个汉字）！", vbInformation, gstrSysName
        Me.txt五笔.SetFocus: Exit Function
    End If
    
    If LenB(StrConv(Trim(Me.txt英文.Text), vbFromUnicode)) > txt英文.MaxLength Then
        MsgBox "英文名称长度超长（最多" & txt英文.MaxLength & "个字符或" & txt英文.MaxLength \ 2 & "个汉字）！", vbInformation, gstrSysName
        Me.txt英文.SetFocus: Exit Function
    End If
    If Trim(Me.cbo单位.Text) = "" Then
        MsgBox "请输入散装单位！", vbInformation, gstrSysName
        Me.cbo单位.SetFocus: Exit Function
    End If
    If zlClinicCodeRepeat(txt编码.Text, Val(mstr诊疗ID)) = True Then
        Me.txt编码.SetFocus: Exit Function
    End If

    '别名检查
    strTemp = ";" & Trim(Me.txt名称.Text) & ";" & Trim(Me.txt英文.Text)
    With Me.vsEditBill
        For i = 1 To .Rows - 1
            If Trim(.TextMatrix(i, .ColIndex("材料名称"))) <> "" Then
                If InStr(1, strTemp & ";", ";" & Trim(.TextMatrix(i, .ColIndex("材料名称"))) & ";") > 0 Then
                    MsgBox "别名存在重复（包括通用名称和英文名称）！", vbInformation, gstrSysName
                    .Row = i: .Col = .ColIndex("材料名称")
                    .SetFocus: Exit Function
                Else
                    strTemp = strTemp & ";" & Trim(.TextMatrix(i, .ColIndex("材料名称")))
                End If
            End If
            If zlCommFun.ActualLen(.TextMatrix(i, .ColIndex("材料名称"))) > txt英文.MaxLength Then
                MsgBox "别名中最多能输入" & txt英文.MaxLength & "个字符或 " & txt英文.MaxLength \ 2 & "个字汉字,请检查！", vbInformation, gstrSysName
                .Row = i: .Col = .ColIndex("材料名称")
                .SetFocus: Exit Function
            End If
            If InStr(1, .TextMatrix(i, .ColIndex("材料名称")), "|") > 0 Then
                MsgBox "别名中不能包含字符“|”！", vbInformation, gstrSysName
                .Row = i: .Col = .ColIndex("材料名称")
                .SetFocus: Exit Function
            End If
            If zlCommFun.ActualLen(.TextMatrix(i, .ColIndex("五笔码"))) > txt五笔.MaxLength Then
                MsgBox "五笔码中最多能输入" & txt五笔.MaxLength & "个字符或 " & txt五笔.MaxLength \ 2 & "个字汉字,请检查！", vbInformation, gstrSysName
                .Row = i: .Col = .ColIndex("五笔码")
                .SetFocus: Exit Function
            End If
            If zlCommFun.ActualLen(.TextMatrix(i, .ColIndex("拼音码"))) > txt五笔.MaxLength Then
                MsgBox "拼音码中最多能输入" & txt五笔.MaxLength & "个字符或 " & txt五笔.MaxLength \ 2 & "个字汉字,请检查！", vbInformation, gstrSysName
                .Row = i: .Col = .ColIndex("拼音码")
                .SetFocus: Exit Function
            End If
            
            If InStr(1, .TextMatrix(i, .ColIndex("五笔码")), "|") > 0 Then
                MsgBox "五笔码中不能包含字符“|”！", vbInformation, gstrSysName
                .Row = i: .Col = .ColIndex("五笔码")
                .SetFocus: Exit Function
            End If
            If InStr(1, .TextMatrix(i, .ColIndex("拼音码")), "|") > 0 Then
                MsgBox "拼音码中不能包含字符“|”！", vbInformation, gstrSysName
                .Row = i: .Col = .ColIndex("拼音码")
                .SetFocus: Exit Function
            End If
        Next
    End With
    ISValied = True
End Function

Public Function zlClinicCodeRepeat(str编码 As String, Optional lngSelfID As Long) As Boolean
    '------------------------------------------------------------------------------------------------------
    '功能：检查诊疗项目编码的是否与现有编码重复，重复则给出提示
    '入参：strInputCode-输入的编码；lngSelfID-自己的ID号，当修改时，需要将自身除开才能判断
    '出参：重复返回True；否则反馈Flase
    '编制:刘兴宏
    '日期:2007/05/24
    '------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    
    strSql = "select K.名称||' ['||I.编码||']'||I.名称 as 名称" & _
            " from 诊疗项目目录 I,诊疗项目类别 K" & _
            " where I.类别=K.编码 and I.编码=[1] " & _
            "       and I.ID<>[2]"
    err = 0: On Error GoTo ErrHand
        
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "检查是否存在重复的编码", str编码, lngSelfID)
        
    With rsTmp
        If .RecordCount <> 0 Then
            MsgBox "该项目与“" & !名称 & "”编码重复！", vbExclamation, gstrSysName
            zlClinicCodeRepeat = True
        Else
            zlClinicCodeRepeat = False
        End If
    End With
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    zlClinicCodeRepeat = True
End Function

Private Function SaveData() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:保存卫生材料品种数据
    '--入参数:
    '--出参数:
    '--返  回:保存成功,返回true,否则返回false
    '-----------------------------------------------------------------------------------------------------------
    Dim lng诊疗ID As Long, intTemp As Integer, i As Long
    Dim strTemp As String
    Dim str站点 As String
    
    If mintEditType = g新增 Then
        lng诊疗ID = sys.NextId("诊疗项目目录")
        gstrSQL = "zl_材料品种_INSERT("
        Me.cmdOK.Tag = lng诊疗ID
        mlng品种id = lng诊疗ID
    Else
        lng诊疗ID = Val(mstr诊疗ID)
        gstrSQL = "zl_材料品种_UPDATE("
        Me.cmdOK.Tag = lng诊疗ID
    End If
    
    strTemp = ""
    With Me.vsEditBill
        For i = 1 To .Rows - 1
            If Trim(.TextMatrix(i, .ColIndex("材料名称"))) <> "" Then
                strTemp = strTemp & "|" & Trim(.TextMatrix(i, .ColIndex("材料名称")))
                strTemp = strTemp & "^" & Trim(.TextMatrix(i, .ColIndex("拼音码")))
                strTemp = strTemp & "^" & Trim(.TextMatrix(i, .ColIndex("五笔码")))
            End If
        Next
    End With
    If strTemp <> "" Then strTemp = Mid(strTemp, 2)
    '检查别名串长度
    If LenB(strTemp) > 4000 Then
        vsEditBill.SetFocus
        MsgBox "别名字符串太长，请减少别名个数或者别名长度。", vbExclamation, gstrSysName
        Exit Function
    End If
    
    If cmbStationNo.Text = "" Then
        str站点 = "Null"
    Else
        str站点 = Mid(cmbStationNo.Text, 1, InStr(1, cmbStationNo.Text, "-") - 1)
    End If
    
    'Zl_材料品种_Update Or zl_材料品种_INSERT
    '  分类id_In In 诊疗项目目录.分类id%Type := Null,
    '  Id_In     In 诊疗项目目录.ID%Type,
    '  编码_In   In 诊疗项目目录.编码%Type,
    '  名称_In   In 诊疗项目目录.名称%Type,
    '  单位_In   In 诊疗项目目录.计算单位%Type := Null,
    '  拼音_In   In 诊疗项目别名.简码%Type := Null,
    '  五笔_In   In 诊疗项目别名.简码%Type := Null,
    '  英文_In   In 诊疗项目别名.名称%Type := Null,
    '  站点_In   In 诊疗项目目录.站点%Type := Null,
    '  别名_In   In Varchar2 := Null --以"|"分隔的别名记录，每条记录按"名称^拼音^五笔"组织
    gstrSQL = gstrSQL & "" & mlng分类id & ","
    gstrSQL = gstrSQL & "" & lng诊疗ID & ","
    gstrSQL = gstrSQL & "'" & Trim(Me.txt编码.Text) & "',"
    gstrSQL = gstrSQL & "'" & Trim(Me.txt名称.Text) & "',"
    gstrSQL = gstrSQL & "'" & Trim(Me.cbo单位.Text) & "',"
    gstrSQL = gstrSQL & "'" & Trim(Me.txt拼音.Text) & "',"
    gstrSQL = gstrSQL & "'" & Trim(Me.txt五笔.Text) & "',"
    gstrSQL = gstrSQL & "'" & Trim(Me.txt英文.Text) & "',"
    gstrSQL = gstrSQL & IIf(cmbStationNo.Visible = True And Trim(cmbStationNo.Text) <> "", "'" & str站点 & "'", "NULL") & ","
    gstrSQL = gstrSQL & "" & Left(Me.cbo适用性别.Text, 1) & ","
    gstrSQL = gstrSQL & "'" & strTemp & "')"
    
    err = 0: On Error GoTo ErrHand
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    SaveData = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub cmdOK_Click()
    Dim intTemp As Integer
    '检查规格页面的输入项是否正确
    If ISValied = False Then Exit Sub
    If mintEditType <> g新增 And mintEditType <> g修改 Then
        Unload Me
        Exit Sub
    End If
    
    If SaveData = False Then Exit Sub
    
    mintSuccess = mintSuccess + 1
    
    If mintEditType = g新增 Then
'        intTemp = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\卫材增加模式\", "品种->规格", "0"))
'        intTemp = Val(zlDatabase.GetPara("品种规格模式", glngSys, mlngModule, "0"))
'        If intTemp = 1 Then
'            '需要增加规格
'            Call frmStuffSpec.ShowEditCard(Me, g新增, Val(Me.cmdOK.Tag), "", mstrPrivs)
'        End If
''        intTemp = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\卫材增加模式\", "品种", "0"))
'        intTemp = Val(zlDatabase.GetPara("品种增加模式", glngSys, mlngModule, "0"))
'        If intTemp = 1 Then
'            Call InitCardData(0)
'            If txt名称.Enabled Then Me.txt名称.SetFocus
'        Else
'            Unload Me
'            Exit Sub
'        End If
        Select Case ActiveControl
            Case cmdSaveAddItem '连续增加品种
                Call InitCardData(0)
                If txt名称.Enabled Then Me.txt名称.SetFocus
            Case cmdSaveAddSpec '连续增加规格
                mlng品种id = Val(Me.cmdOK.Tag)
                Unload Me
                Call frmStuffSpec.ShowEditCard(frmStuffMgr, g新增, mlng品种id, mlng分类id, "", mstrPrivs)
            Case Else   '直接保存退出
                Unload Me
        End Select
    Else
        Unload Me
    End If
    
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Sub cmdCancel_Click()
    Unload Me
    Exit Sub
End Sub
Private Sub cmd帮助_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int(glngSys / 100))
End Sub

Private Function GetMaxCode() As String
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:获取最大编号
    '--入参数:
    '--出参数:
    '--返  回:最大编号
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim rsCode As ADODB.Recordset
    Dim strTemp As String
    Dim intCodeType As Integer
    Dim str编码 As String
    
    On Error GoTo ErrHandle
    intCodeType = Val(zlDatabase.GetPara("编码递增模式", glngSys, mlngModule))
    strTemp = Mid(Me.txt分类.Text, 2, InStr(1, Me.txt分类.Text, "]") - 2)
    
    If intCodeType = 0 Or Len(strTemp) >= 16 Then
    '0000000001、0000000002
        gstrSQL = "Select Nvl(编码, '000000000') As 编码" & vbNewLine & _
                        "From (Select 编码 From 诊疗项目目录 Where 类别 = '4' Order By Length(编码) Desc, 编码 Desc)" & vbNewLine & _
                        "Where Rownum = 1"

        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        With rsTemp
            str编码 = zlCommFun.IncStr(!编码)
            GetMaxCode = str编码
        End With
      
    Else
    
        gstrSQL = "Select a.Id, a.分类id, a.编码, a.名称 From 诊疗项目目录 A Where 分类id =[1] Order By 分类id, ID"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng分类id)

        If Len(strTemp) >= 7 Then
            str编码 = "01"
            str编码 = IIf(intCodeType = 1, "4", "") & strTemp & str编码
        Else
            str编码 = Mid("000000000", 1, 9 - Len(strTemp) - IIf(intCodeType = 1, 1, 0))
            str编码 = IIf(intCodeType = 1, "4", "") & strTemp & str编码
            str编码 = zlCommFun.IncStr(str编码)
        End If
        
        GetMaxCode = str编码
    
        Do While True
            rsTemp.Filter = ""
            rsTemp.Filter = "编码='" & GetMaxCode & "'"
            If rsTemp.RecordCount = 0 Then
                Exit Do
            End If
            GetMaxCode = zlCommFun.IncStr(GetMaxCode)
    
            rsTemp.MoveNext
        Loop
    End If
    
    gstrSQL = "Select 编码 From 诊疗项目目录 "
    Set rsCode = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        
    Do While True
        rsCode.Filter = ""
        rsCode.Filter = "编码='" & GetMaxCode & "'"
        If rsCode.RecordCount = 0 Then
            Exit Do
        End If
        GetMaxCode = zlCommFun.IncStr(GetMaxCode)
    Loop
    
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub Load诊疗分类信息()
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:加载选择
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim objNode As Node
    
    On Error GoTo ErrHandle
    '分类选择树装入
    gstrSQL = "select ID,上级ID,编码,名称,简码" & _
            " From 诊疗分类目录" & _
            " Where 类型 =7 " & _
            " start with 上级ID is null" & _
            " connect by prior ID=上级ID"
    zlDatabase.OpenRecordset rsTemp, gstrSQL, Me.Caption
    
    With rsTemp
        Me.tvwClass.Nodes.Clear
        Do While Not .EOF
            If IsNull(!上级ID) Then
                Set objNode = Me.tvwClass.Nodes.Add(, , "_" & !Id, "[" & !编码 & "]" & !名称, "close")
            Else
                Set objNode = Me.tvwClass.Nodes.Add("_" & !上级ID, tvwChild, "_" & !Id, "[" & !编码 & "]" & !名称, "close")
            End If
            objNode.Sorted = True
            objNode.Tag = IIf(IsNull(!简码), "", !简码)
            objNode.ExpandedImage = "expend"
            .MoveNext
        Loop
        err = 0: On Error Resume Next
        Me.tvwClass.Nodes("_" & mlng分类id).Selected = True
        Me.txt分类.Text = Me.tvwClass.SelectedItem.Text
        mlng分类id = Mid(Me.tvwClass.SelectedItem.Key, 2)
    End With

    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
    
End Sub

Private Sub Form_Load()
    Dim strSql As String
    Dim rsrecord As ADODB.Recordset
    
    On Error GoTo ErrHandle
    mblnFrist = True
'    With cmbStationNo
'        .Clear
'        .AddItem ""
'        .AddItem "0"
'        .AddItem "1"
'        .AddItem "2"
'        .AddItem "3"
'        .AddItem "4"
'        .AddItem "5"
'        .AddItem "6"
'        .AddItem "7"
'        .AddItem "8"
'        .AddItem "9"
'        .ListIndex = 0
'    End With
    strSql = "select 编号,名称 from zlnodelist"
    Set rsrecord = zlDatabase.OpenSQLRecord(strSql, "站点查询")
    With cmbStationNo
        .AddItem ""
        Do While Not rsrecord.EOF
            .AddItem rsrecord!编号 & "-" & rsrecord!名称
            rsrecord.MoveNext
        Loop
    End With
    If mintEditType <> g新增 Then cmdSaveAddItem.Enabled = False: cmdSaveAddSpec.Enabled = False
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If tvwClass.Visible Then
        tvwClass.Visible = False
        Cancel = 1
        Exit Sub
    End If
End Sub

Private Sub txt编码_Change()
    mblnChange = True
End Sub

Private Sub txt编码_GotFocus()
    ImeLanguage False
    zlControl.TxtSelAll txt编码
End Sub

Private Sub txt编码_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub txt编码_KeyPress(KeyAscii As Integer)
        Select Case KeyAscii
        Case vbKeyBack, vbKeyEscape, 3, 22
            Exit Sub
        Case Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
            If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
        End Select
        KeyAscii = 0
End Sub

Private Sub txt编码_LostFocus()
    ImeLanguage False
End Sub

Private Sub txt分类_Change()
    mlng分类id = 0
    txt分类.Tag = ""
End Sub

Private Sub txt分类_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub
Private Sub tvwClass_DblClick()
    If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
    Me.txt分类.Text = Me.tvwClass.SelectedItem.Text
    mlng分类id = Val(Mid(Me.tvwClass.SelectedItem.Key, 2))
    txt分类.Tag = mlng分类id
    txt编码.Text = GetMaxCode
    If txt名称.Enabled Then Me.txt名称.SetFocus
End Sub

Private Sub tvwClass_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn
        If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
        If Me.tvwClass.SelectedItem.Children > 0 Then Exit Sub
        Call tvwClass_DblClick
    Case vbKeySpace
        Call tvwClass_DblClick
    Case vbKeyEscape
        Call tvwClass_LostFocus
    End Select
End Sub

Private Sub tvwClass_LostFocus()
    If Me.cmd分类 Is ActiveControl Then
        Exit Sub
    End If
    Me.tvwClass.Visible = False
End Sub

Private Sub txt名称_Change()
    mblnChange = True
    '拼音和五笔
    Me.txt拼音.Text = zlStr.GetCodeByORCL(Me.txt名称.Text, 0, Me.txt拼音.MaxLength)
    Me.txt五笔.Text = zlStr.GetCodeByORCL(Me.txt名称.Text, 1, Me.txt五笔.MaxLength)
End Sub

Private Sub txt名称_GotFocus()
    ImeLanguage True
    zlControl.TxtSelAll txt名称
End Sub

Private Sub Txt名称_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub txt名称_LostFocus()
    ImeLanguage False
End Sub

Private Sub txt拼音_Change()
    mblnChange = True
End Sub

Private Sub txt拼音_GotFocus()
    ImeLanguage False
End Sub

Private Sub txt拼音_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If

End Sub

Private Sub txt五笔_Change()
    mblnChange = True
    
End Sub

Private Sub txt五笔_GotFocus()
    ImeLanguage False
End Sub

Private Sub txt五笔_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If

End Sub

Private Sub txt英文_Change()
    mblnChange = True
End Sub

Private Sub txt英文_GotFocus()
    ImeLanguage False
    zlControl.TxtSelAll txt英文
End Sub

Private Sub txt英文_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub

 
Private Sub vsEditBill_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim strKey As String
    With vsEditBill
        Select Case Col
        Case .ColIndex("材料名称")
            strKey = Trim(.TextMatrix(.Row, .Col))
            If strKey = "" Then Exit Sub
            .TextMatrix(Row, .ColIndex("拼音码")) = zlStr.GetCodeByORCL(strKey, 0, Me.txt拼音.MaxLength)
            .TextMatrix(Row, .ColIndex("五笔码")) = zlStr.GetCodeByORCL(strKey, 1, Me.txt五笔.MaxLength)
            If .Row = .Rows - 1 Then
                .Rows = .Rows + 1
                .Row = .Rows - 1
            End If
        Case .ColIndex("拼音码")
            If Trim(.TextMatrix(.Row, .ColIndex("材料名称"))) = "" Then Exit Sub
            If .Row = .Rows - 1 Then
                .Rows = .Rows + 1: .Row = .Rows - 1
            End If
        Case Else
        End Select
    End With
    '重算行号
    Call RedoRowNo
End Sub

Private Sub vsEditBill_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsEditBill
        Select Case Col
        Case .ColIndex("材料名称")
        Case .ColIndex("五笔码"), .ColIndex("拼音码")
            If Trim(.Cell(flexcpData, Row, .ColIndex("材料名称"))) = "" Then
                Cancel = True
            End If
        Case Else
            Cancel = True
        End Select
    End With
End Sub

Private Sub vsEditBill_CellChanged(ByVal Row As Long, ByVal Col As Long)
    mblnChange = True
End Sub

Private Sub vsEditBill_EnterCell()
    If mintEditType = g查看 Then Exit Sub
    If vsEditBill.Col = vsEditBill.ColIndex("材料名称") Then
        OS.OpenIme True
        vsEditBill.EditMaxLength = Me.txt名称.MaxLength
    Else
        OS.OpenIme False
    End If
End Sub

Private Sub vsEditBill_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngCol As Long
    If mintEditType <> g查看 Then
        With vsEditBill
            If KeyCode = vbKeyDelete Then
                If MsgBox("你是否真的要删除该行的材料别名吗?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
                If .Row = .Rows - 1 And .Row = 1 Then
                    For lngCol = 0 To .Cols - 1
                        .TextMatrix(.Row, lngCol) = ""
                        .Cell(flexcpData, .Row, lngCol) = ""
                    Next
                Else
                    .RemoveItem .Row
                End If
            End If
            Call RedoRowNo
        End With
    End If
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    With vsEditBill
        If Trim(.TextMatrix(.Row, .ColIndex("材料名称"))) = "" Then
            OS.PressKey vbKeyTab
            Exit Sub
        End If
        Select Case .Col
        Case .ColIndex("五笔码")
            .Col = .ColIndex("材料名称")
            If .Row >= .Rows - 1 Then
                If mintEditType = g查看 Then
                Else
                    .Rows = .Rows + 1
                End If
                .Row = .Rows - 1
            Else
                .Row = .Row + 1
            End If
            .SetFocus
        Case Else
            OS.PressKey vbKeyRight
        End Select
    End With
End Sub

Private Sub vsEditBill_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub vsEditBill_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0: Exit Sub
    If KeyAscii = Asc("|") Then KeyAscii = 0: Exit Sub
    If KeyAscii = Asc("^") Then KeyAscii = 0: Exit Sub
    
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        If Col < vsEditBill.ColIndex("五笔码") Then
            If Col = vsEditBill.ColIndex("材料名称") Then
                OS.PressKey vbKeyDown
            Else
                OS.PressKey vbKeyRight
            End If
        End If
        Exit Sub
    End If
    
    With vsEditBill
        Select Case Col
        Case .ColIndex("材料名称"), .ColIndex("拼音码"), .ColIndex("五笔码")
            Call VsFlxGridCheckKeyPress(vsEditBill, Row, Col, KeyAscii, m文本式)
        Case Else
        End Select
    End With
End Sub

Private Sub vsEditBill_LeaveCell()
    If mintEditType = g查看 Then Exit Sub
    OS.OpenIme False
End Sub

Private Sub vsEditBill_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
  Dim strKey As String
    Dim intCol As Integer
    Dim strTemp As String
    
    If mintEditType = g查看 Then Cancel = True: Exit Sub
    
    strKey = Trim(vsEditBill.EditText)
    strKey = Replace(strKey, Chr(vbKeyReturn), "")
    strKey = Replace(strKey, Chr(10), "")
    With vsEditBill
        Select Case Col
        Case .ColIndex("材料名称")
            If zlCommFun.ActualLen(strKey) > txt英文.MaxLength Then
                ShowMsgBox "材料名称必须为小于等于" & txt英文.MaxLength & "个字符或" & Int(txt英文.MaxLength / 2) & "个汉字,请重新输入！"
                Cancel = True
                Exit Sub
            End If
        Case .ColIndex("五笔码"), .ColIndex("拼音码") '
            If zlCommFun.ActualLen(strKey) > txt五笔.MaxLength Then
                ShowMsgBox vsEditBill.TextMatrix(0, Col) & "必须为小于等于" & txt五笔.MaxLength & "个字符或" & txt五笔.MaxLength \ 2 & "个汉字,请重新输入！"
                Cancel = True
                Exit Sub
            End If
        End Select
    End With
End Sub
Private Sub RedoRowNo()
    '------------------------------------------------------------------------------
    '功能:重置行号
    '返回:
    '编制:刘兴宏
    '日期:2007/08/14
    '------------------------------------------------------------------------------
    Dim i As Long
    With vsEditBill
        For i = 1 To .Rows - 1
            .TextMatrix(i, .ColIndex("行号")) = i
        Next
    End With

End Sub

Private Sub cmbStationNo_Change()
    mblnChange = True
End Sub

Private Sub cmbStationNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    OS.PressKey vbKeyTab
    
End Sub
