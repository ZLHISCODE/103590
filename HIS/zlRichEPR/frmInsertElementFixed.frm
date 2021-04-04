VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmInsertElementFixed 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "选择所见项"
   ClientHeight    =   5730
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   7140
   ControlBox      =   0   'False
   Icon            =   "frmInsertElementFixed.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   7140
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picFind 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   90
      ScaleHeight     =   330
      ScaleWidth      =   2085
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   180
      Width           =   2085
      Begin VB.CommandButton cmdFind 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1620
         Picture         =   "frmInsertElementFixed.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   0
         Width           =   300
      End
      Begin VB.TextBox txtFind 
         Height          =   300
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   1590
      End
   End
   Begin VB.PictureBox picRight 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4515
      Left            =   2745
      ScaleHeight     =   4515
      ScaleWidth      =   3615
      TabIndex        =   3
      Top             =   315
      Width           =   3615
      Begin VB.PictureBox picExample 
         BorderStyle     =   0  'None
         Height          =   780
         Left            =   180
         ScaleHeight     =   780
         ScaleWidth      =   3030
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   3015
         Width           =   3030
         Begin RichTextLib.RichTextBox rtbThis 
            Height          =   420
            Left            =   135
            TabIndex        =   5
            Top             =   180
            Width           =   2310
            _ExtentX        =   4075
            _ExtentY        =   741
            _Version        =   393217
            BackColor       =   -2147483633
            BorderStyle     =   0
            Enabled         =   -1  'True
            ReadOnly        =   -1  'True
            ScrollBars      =   3
            Appearance      =   0
            TextRTF         =   $"frmInsertElementFixed.frx":0396
         End
         Begin VB.Shape shpExample 
            BorderColor     =   &H00808080&
            Height          =   375
            Left            =   90
            Top             =   135
            Width           =   330
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vfgThis 
         Height          =   2595
         Left            =   180
         TabIndex        =   6
         Top             =   180
         Width           =   3000
         _cx             =   5292
         _cy             =   4577
         Appearance      =   0
         BorderStyle     =   0
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
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
         GridColor       =   15198183
         GridColorFixed  =   15198183
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   0
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   240
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmInsertElementFixed.frx":0433
         ScrollTrack     =   -1  'True
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
         Ellipsis        =   1
         ExplorerBar     =   7
         PicturesOver    =   -1  'True
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
         WallPaperAlignment=   1
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Shape shpFG 
         BorderColor     =   &H00808080&
         Height          =   375
         Left            =   45
         Top             =   45
         Width           =   330
      End
   End
   Begin VB.Frame fraList 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4245
      Left            =   45
      TabIndex        =   0
      Top             =   630
      Width           =   2160
      Begin VB.CommandButton cmdTab 
         Caption         =   "标记图"
         Height          =   300
         Index           =   0
         Left            =   0
         TabIndex        =   1
         Top             =   2040
         Visible         =   0   'False
         Width           =   1335
      End
      Begin MSComctlLib.TreeView tvwItem 
         Height          =   1995
         Index           =   0
         Left            =   15
         TabIndex        =   2
         Top             =   0
         Visible         =   0   'False
         Width           =   4860
         _ExtentX        =   8573
         _ExtentY        =   3519
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   494
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         FullRowSelect   =   -1  'True
         ImageList       =   "iLsTree"
         Appearance      =   1
      End
   End
   Begin MSComctlLib.ImageList iLsTree 
      Left            =   1575
      Top             =   4995
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInsertElementFixed.frx":04B3
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInsertElementFixed.frx":060D
            Key             =   "Attr"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInsertElementFixed.frx":0927
            Key             =   "Class"
         EndProperty
      EndProperty
   End
   Begin VB.Image imgY 
      Height          =   5115
      Left            =   2295
      MouseIcon       =   "frmInsertElementFixed.frx":1201
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   45
   End
End
Attribute VB_Name = "frmInsertElementFixed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private clsComLib As New zl9ComLib.clsComLib

Private iCurrTab As Integer
Private frmParent As Object

Public ItemID As String

Public Sub ShowMe(Parent As Object)
    Set frmParent = Parent
    Me.Show , Parent
End Sub

Private Sub cmdFind_Click()
    Dim strFind As String
    strFind = UCase(Trim(txtFind))
    Dim i As Long
    For i = 1 To tvwItem(iCurrTab).Nodes.Count
        If tvwItem(iCurrTab).Nodes(i).Tag Like strFind & "*" Then
            tvwItem(iCurrTab).Nodes(i).Selected = True
            tvwItem_NodeClick iCurrTab, tvwItem(iCurrTab).SelectedItem
            Exit For
        End If
    Next
End Sub

Private Sub cmdTab_Click(Index As Integer)
    If iCurrTab = Index Then Me.tvwItem(Index).SetFocus: Exit Sub
    iCurrTab = Index
        
    Form_Resize
    If tvwItem(iCurrTab).Nodes.Count > 0 Then
        Set tvwItem(iCurrTab).SelectedItem = tvwItem(iCurrTab).Nodes(1)
        tvwItem(iCurrTab).SetFocus
        tvwItem_NodeClick iCurrTab, tvwItem(iCurrTab).SelectedItem
    Else
        vfgThis.ListItems.Clear
    End If
End Sub

Private Sub vfgThis_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF7 Then
        Me.Hide
        KeyCode = 0
    End If
End Sub

Private Sub vfgThis_RowColChange()
    Dim lngID As Long
    If vfgThis.Row > 0 And vfgThis.Rows > 1 Then
        lngID = Val(vfgThis.Cell(flexcpText, vfgThis.Row, 0))
        FillExample lngID
    End If
End Sub

Private Sub FillExample(lngID As Long)
    '填充示例文本
    Dim i As Long, j As Long, lngLen As Long
    LockWindowUpdate rtbThis.hWnd
    Dim Rs As New ADODB.Recordset
    Rs.CursorLocation = adUseClient
    Set Rs = OpenSQLRecord("select * FROM 诊治所见项目 where ID=[1]", Me.Caption, lngID)
    With rtbThis
        If Rs.EOF Then
            .Text = ""
        Else
            .Text = Rs("中文名") & "：" & IIf(Trim(NVL(Rs("初始值"))) = "", "  ", NVL(Rs("初始值"))) & Rs("单位")
            lngLen = Len(Rs("中文名"))
            .SelStart = lngLen + 1
            .SelLength = Len(IIf(Trim(NVL(Rs("初始值"))) = "", "  ", NVL(Rs("初始值"))))
            .SelColor = vbBlue
            .SelUnderline = True
        End If
    End With
    LockWindowUpdate 0
End Sub

Private Sub Form_Activate()
    ItemID = ""
    txtFind.SelStart = 0
    txtFind.SelLength = Len(txtFind)
    txtFind.SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Me.Hide
    ElseIf KeyAscii = vbKeyReturn Then
        If ActiveControl.Name = "txtFind" Then
            If vfgThis.Rows > 1 Then vfgThis.Row = 1
            vfgThis.SetFocus
        ElseIf ActiveControl.Name = "vfgThis" Then
            Call vfgThis_DblClick
        End If
    ElseIf KeyAscii <> vbKeyBack And ActiveControl.Name <> "txtFind" Then
        txtFind.Text = txtFind.Text + Chr(KeyAscii)
        txtFind.SelStart = Len(txtFind)
        txtFind.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Me.Left = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "MainLeft", (Screen.Width - 6000) / 2)
    Me.Top = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "MainTop", (Screen.Height - 5000) / 2)
    Me.Width = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "MainWidth", 6000)
    Me.Height = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "MainHeight", 5000)
    
    CreateItemTree
    
    On Error Resume Next
    iCurrTab = 1
    Set tvwItem(1).SelectedItem = tvwItem(1).Nodes(1)
    tvwItem_NodeClick iCurrTab, tvwItem(iCurrTab).SelectedItem
    
    vfgThis.Editable = flexEDNone
    Me.KeyPreview = True
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    With picRight
        .Left = imgY.Left + imgY.Width + Screen.TwipsPerPixelX: .Top = Screen.TwipsPerPixelY * 2
        .Width = Me.ScaleWidth - .Left - Screen.TwipsPerPixelX * 2: .Height = Me.ScaleHeight - Screen.TwipsPerPixelY * 4
        .Refresh
    End With
    
    '   显示选项卡
    ShowList imgY.Left
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '保存窗体位置
    If Me.WindowState <> vbMinimized Then
        SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "MainLeft", Me.Left
        SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "MainTop", Me.Top
        SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "MainWidth", Me.Width
        SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "MainHeight", Me.Height
    End If
End Sub

Private Sub imgY_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Not Button = vbLeftButton Then Exit Sub
            
    imgY.Left = imgY.Left + X
    If imgY.Left < 1000 Then imgY.Left = 1000
    If imgY.Left > 3000 Then imgY.Left = 3000
    
    Form_Resize
End Sub

'创建所见项分类及其项目的TreeView
Private Sub CreateItemTree()
    Dim rsItem As New ADODB.Recordset
    Dim sCurID As String
    Dim iStackPoint As Integer '堆栈指针
    Dim aStack() As String '堆栈
    Dim TmpNode As Node
    Dim i As Integer, AttrID As String
    
    '从诊治所见性质中提取
    
    Set rsItem = OpenSQLRecord("Select * From 诊治所见性质 Order By 编码", Me.Caption)
    Do While Not rsItem.EOF
        Load cmdTab(cmdTab.Count)
        With cmdTab(cmdTab.Count - 1)
            .Caption = rsItem("名称") '+ IIf(rsItem("固定") = 1, "（只读）", "")
            .Tag = rsItem("固定") & "-" & rsItem("编码")
            .ZOrder 0
            .Visible = True
        End With
        Load tvwItem(tvwItem.Count)
        tvwItem(tvwItem.Count - 1).Visible = True
        
        rsItem.MoveNext
    Loop
    
    For i = 1 To cmdTab.Count - 1
        AttrID = Mid(cmdTab(i).Tag, InStr(cmdTab(i).Tag, "-") + 1)
    
        ReDim aStack(0)
        aStack(0) = ""
        iStackPoint = 0
        
        Do While iStackPoint > -1
            sCurID = aStack(iStackPoint)
            '添加下级所见项分类
            Set rsItem = OpenSQLRecord("Select * From 诊治所见分类 Where 上级ID" + IIf(sCurID = "", " is null ", "='" + sCurID + "' ") + "And 性质='" + AttrID + "'", Me.Caption)
            
            '该分类的下级已处理，将其从堆栈中弹出
            iStackPoint = iStackPoint - 1
            
            Do While Not rsItem.EOF
                If sCurID = "" Then
                    Set TmpNode = tvwItem(i).Nodes.Add(, , "Key" & rsItem("ID"), rsItem("名称"), "Class")
                Else
                    Set TmpNode = tvwItem(i).Nodes.Add("Key" + sCurID, tvwChild, "Key" & rsItem("ID"), rsItem("名称"), "Class")
                End If
                TmpNode.Tag = rsItem("简码") & "||" & rsItem("性质") & "||" & rsItem("编码") & "||" & rsItem("名称")
                
                '将新分类压入堆栈
                iStackPoint = iStackPoint + 1
                ReDim Preserve aStack(iStackPoint)
                aStack(iStackPoint) = rsItem("ID")
                
                rsItem.MoveNext
            Loop
        Loop
    Next
End Sub

Private Sub ShowSubItem(ByVal NodeID As String, ByVal AttributeID As String)
    Dim Rs As New ADODB.Recordset, i As Long
    Dim sSQL As String
    
    vfgThis.Clear
    vfgThis.Rows = 1
    vfgThis.Cell(flexcpText, 0, 2) = "编码"
    vfgThis.Cell(flexcpText, 0, 3) = "名称"
    vfgThis.ColAlignment(2) = flexAlignLeftCenter
    
    '添加下级所见项目
    sSQL = "Select ID,编码,中文名,nvl(英文名,' '),nvl(替换域,1),nvl(类型,0)," + _
       "nvl(长度,10),nvl(小数,0),nvl(单位,' '),nvl(表示法,0),nvl(性别域,0)," + _
       "nvl(数值域,' '),nvl(正常域,' '),nvl(初始值,' '),nvl(文字表述,1),nvl(空值文字,' '),nvl(临床意义,' ') " + _
       "From 诊治所见项目 Where " + IIf(NodeID = "", "性质='" + AttributeID + "' And 分类ID is null ", "分类ID='" + NodeID + "' ")
    Set Rs = OpenSQLRecord(sSQL, Me.Caption)
    
    i = 1
    If Not Rs.EOF Then
        vfgThis.Cols = 4
        vfgThis.Rows = Rs.RecordCount + 1
        Do While Not Rs.EOF
            vfgThis.Cell(flexcpText, i, 0) = Rs("ID")
            vfgThis.Cell(flexcpPicture, i, 1) = iLsTree.ListImages("Item").Picture
            vfgThis.Cell(flexcpText, i, 2) = Rs("编码")
            vfgThis.Cell(flexcpText, i, 3) = Rs("中文名")
            i = i + 1
            Rs.MoveNext
        Loop
    End If
End Sub

Private Sub vfgThis_DblClick()
    '插入诊治要素
    
    Dim lngID As Long, strTMP As String
    Dim i As Long, j As Long, p As Long, q As Long

    '查找关键字 ID ！
    Dim sType As String, lSS As Long, lSE As Long, lES As Long, lEE As Long, lID As Long, bInKeys As Boolean, bNeeded As Boolean
    bInKeys = frmParent.IsBetweenAnyKeys(frmParent.Editor1.SelStart + 1, sType, lSS, lSE, lES, lEE, lID, bNeeded)
    If bInKeys Then Exit Sub

    '双击在文档中插入元素（所见项目）
    With frmParent.Editor1
        If vfgThis.Row > 0 And vfgThis.Rows > 1 Then
            lngID = Val(vfgThis.Cell(flexcpText, vfgThis.Row, 0))
            '插入元素
            Dim Rs As New ADODB.Recordset
            Rs.CursorLocation = adUseClient
            Set Rs = OpenSQLRecord("select * FROM 诊治所见项目 where ID=[1]", Me.Caption, lngID)
            If Not Rs.EOF Then
                Dim Ele As New cEPRElement
                i = frmParent.Document.Elements.Add
                Set Ele = frmParent.Document.Elements("K" & i)
                Ele.ID = Rs("ID")
                Ele.中文名 = NVL(Rs("中文名"))
                Ele.英文名 = NVL(Rs("英文名"))
                Ele.初始值 = IIf(Trim(NVL(Rs("初始值"))) = "", "  ", NVL(Rs("初始值")))
                Ele.数值域 = NVL(Rs("数值域"))
                Ele.类型 = NVL(Rs("类型"), 0)
                Ele.表示法 = NVL(Rs("表示法"), 0)
                Ele.单位 = NVL(Rs("单位"))
                Ele.替换域 = NVL(Rs("替换域"), 0)

                .Freeze
                .ForceEdit = True

                i = .SelStart
                If Ele.替换域 = 0 Then
                    strTMP = Trim(Ele.中文名) & "："
                    strTMP = strTMP & "ES(" & Format(Ele.流水号, "00000000") & ",1,0)" & Ele.初始值 & Ele.单位
                    strTMP = strTMP & "EE(" & Format(Ele.流水号, "00000000") & ",1,0)" & "，"
                    '字符串整体赋值
                    .Range(i, i).Text = strTMP
    
                    '逐段设置其显示属性
                    i = i + Len(Trim(Ele.中文名) & "：")
                    .Range(i, i + 32 + Len(Ele.初始值) + Len(Ele.单位)).Font.Protected = True
                    .Range(i, i + 16).Font.Hidden = True
                    .Range(i + 16 + Len(Ele.初始值) + Len(Ele.单位), i + 16 + Len(Ele.初始值) + Len(Ele.单位) + 16).Font.Hidden = True
                    .Range(i + 16, i + 16 + Len(Ele.初始值)).Font.Underline = tomHair 'tomthick
                    .Range(i + 16, i + 16 + Len(Ele.初始值)).Font.ForeColor = vbBlue
    
                    .Range(i + 32 + Len(Ele.初始值) + Len(Ele.单位) + 1, i + 32 + Len(Ele.初始值) + Len(Ele.单位) + 1).Selected
                ElseIf Ele.替换域 = 1 Then
                    j = Len("{{" & Ele.英文名 & "}}")
                    strTMP = "ES(" & Format(Ele.流水号, "00000000") & ",1,0){{" & Ele.英文名 & "}}EE(" & Format(Ele.流水号, "00000000") & ",1,0)" & "，"
                    .Range(i, i).Text = strTMP
                    .Range(i, i + j + 32).Font.Protected = True
                    .Range(i, i + 16).Font.Hidden = True
                    .Range(i + 16 + j, i + 32 + j).Font.Hidden = True
                    .Range(i + 32 + j, i + 32 + j + 1).Font.Protected = False
                    .Range(i + 16, i + 16 + j).Font.Underline = tomHair
                    .Range(i + 16, i + 16 + j).Font.ForeColor = vbBlue
                    
                    .Range(i + 32 + j + 1, i + 32 + j + 1).Selected
                End If
                .ForceEdit = False
                .UnFreeze
                .SetFocus
            End If
            Rs.Close
            Set Rs = Nothing
        End If
    End With
    Me.Hide
End Sub

Private Sub picRight_Resize()
    On Error Resume Next
    Dim lngX As Long, lngY As Long
    lngX = Screen.TwipsPerPixelX
    lngY = Screen.TwipsPerPixelY
    
    With picRight
        picExample.Move lngX, .ScaleHeight - picExample.Height, .ScaleWidth - lngX * 2
        vfgThis.Move lngX * 2, lngY * 2, .ScaleWidth - lngX * 4, .ScaleHeight - picExample.Height - lngY * 8
        shpFG.Move lngX, vfgThis.Top - lngY, vfgThis.Width + lngX * 2, vfgThis.Height + lngY * 2
        
        rtbThis.Move lngX * 2, lngY * 2, picExample.ScaleWidth - lngX * 4, picExample.ScaleHeight - 350
        shpExample.Move 0, 0, picExample.Width, picExample.Height
    End With
End Sub

Private Sub tvwItem_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF7 Then
        Me.Hide
        KeyCode = 0
    End If
End Sub

Private Sub tvwItem_NodeClick(Index As Integer, ByVal Node As MSComctlLib.Node)
    If Node Is Nothing Then Exit Sub
    If Node.Key Like "Key_*" Then
        ShowSubItem "", Mid(Node.Key, 5)
    Else
        ShowSubItem Mid(Node.Key, 4), ""
    End If
End Sub

Private Sub ShowList(ByVal Width As Long, Optional ByVal Top As Long = -1)
    Dim i As Integer
    
    picFind.Move 0, 0, Width
    cmdFind.Left = picFind.ScaleWidth - cmdFind.Width
    txtFind.Width = picFind.ScaleWidth - cmdFind.Width
    
    With fraList
        .Left = picFind.Left: .Top = picFind.Height
        .Width = Width
        .Height = Me.ScaleHeight - .Top
        .Visible = True
    End With
    For i = 1 To tvwItem.Count - 1
        tvwItem(i).Visible = IIf(i = iCurrTab, True, False)
        With cmdTab(i)
            If i <= iCurrTab Then
                .Top = (i - 1) * (cmdTab(0).Height - 15)
            Else
                .Top = fraList.Height - (tvwItem.Count - i) * (cmdTab(0).Height - 15)
            End If
            
            .Width = fraList.Width
            .Left = 0
            
            .Visible = True
        End With
    Next
    
    With tvwItem(iCurrTab)
        .Left = picFind.Left
        .Top = cmdTab(iCurrTab).Top + cmdTab(iCurrTab).Height
        .Width = fraList.Width
        .Height = fraList.Height - (tvwItem.Count - iCurrTab - 1) * (cmdTab(0).Height - 15) - .Top
    End With
End Sub

Private Sub txtFind_Change()
    cmdFind_Click
End Sub

Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF7 Then
        Me.Hide
        KeyCode = 0
    End If
End Sub

