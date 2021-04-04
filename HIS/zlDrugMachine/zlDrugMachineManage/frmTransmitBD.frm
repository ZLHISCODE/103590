VERSION 5.00
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmTransmitBD 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "基础数据传送"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8400
   Icon            =   "frmTransmitBD.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   8400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin MSComctlLib.ImageList img24 
      Left            =   1320
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTransmitBD.frx":06EA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img16 
      Left            =   600
      Top             =   120
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
            Picture         =   "frmTransmitBD.frx":0DE4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picSSB 
      BorderStyle     =   0  'None
      Height          =   3975
      Left            =   2280
      ScaleHeight     =   3975
      ScaleWidth      =   5880
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   240
      Width           =   5880
      Begin TabDlg.SSTab sstClass 
         Height          =   3375
         Left            =   120
         TabIndex        =   8
         Top             =   0
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   5953
         _Version        =   393216
         Style           =   1
         Tabs            =   5
         TabsPerRow      =   5
         TabHeight       =   520
         TabCaption(0)   =   "部门信息"
         TabPicture(0)   =   "frmTransmitBD.frx":117E
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "pic(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "人员信息"
         TabPicture(1)   =   "frmTransmitBD.frx":119A
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "pic(1)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "药品目录"
         TabPicture(2)   =   "frmTransmitBD.frx":11B6
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "pic(2)"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "库存信息"
         TabPicture(3)   =   "frmTransmitBD.frx":11D2
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "pic(3)"
         Tab(3).Control(0).Enabled=   0   'False
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "发药窗口"
         TabPicture(4)   =   "frmTransmitBD.frx":11EE
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "pic(4)"
         Tab(4).Control(0).Enabled=   0   'False
         Tab(4).ControlCount=   1
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000003&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   2535
            Index           =   0
            Left            =   120
            ScaleHeight     =   2535
            ScaleWidth      =   4095
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   480
            Width           =   4095
            Begin VB.CheckBox chk 
               Appearance      =   0  'Flat
               BackColor       =   &H80000003&
               Caption         =   "全选"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   0
               Left            =   3240
               TabIndex        =   27
               Top             =   30
               Width           =   735
            End
            Begin MSComctlLib.ListView lvw 
               Height          =   1575
               Index           =   0
               Left            =   120
               TabIndex        =   26
               Top             =   480
               Width           =   3255
               _ExtentX        =   5741
               _ExtentY        =   2778
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   0   'False
               HideColumnHeaders=   -1  'True
               Checkboxes      =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   0
               NumItems        =   0
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "部门性质"
               Height          =   180
               Index           =   0
               Left            =   120
               TabIndex        =   28
               Top             =   75
               Width           =   720
            End
         End
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000003&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   2535
            Index           =   1
            Left            =   -74880
            ScaleHeight     =   2535
            ScaleWidth      =   4095
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   480
            Width           =   4095
            Begin VB.CheckBox chk 
               Appearance      =   0  'Flat
               BackColor       =   &H80000003&
               Caption         =   "全选"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   1
               Left            =   3240
               TabIndex        =   22
               Top             =   30
               Width           =   735
            End
            Begin MSComctlLib.ListView lvw 
               Height          =   1575
               Index           =   1
               Left            =   120
               TabIndex        =   23
               Top             =   480
               Width           =   3255
               _ExtentX        =   5741
               _ExtentY        =   2778
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   0   'False
               HideColumnHeaders=   -1  'True
               Checkboxes      =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   0
               NumItems        =   0
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "工作性质"
               Height          =   180
               Index           =   1
               Left            =   120
               TabIndex        =   24
               Top             =   75
               Width           =   720
            End
         End
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000003&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   2535
            Index           =   2
            Left            =   -74880
            ScaleHeight     =   2535
            ScaleWidth      =   4095
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   480
            Width           =   4095
            Begin VB.CheckBox chk 
               Appearance      =   0  'Flat
               BackColor       =   &H80000003&
               Caption         =   "全选"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   2
               Left            =   3240
               TabIndex        =   18
               Top             =   30
               Width           =   735
            End
            Begin MSComctlLib.ListView lvw 
               Height          =   1575
               Index           =   2
               Left            =   120
               TabIndex        =   19
               Top             =   480
               Width           =   3255
               _ExtentX        =   5741
               _ExtentY        =   2778
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   0   'False
               HideColumnHeaders=   -1  'True
               Checkboxes      =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   0
               NumItems        =   0
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "药品剂型"
               Height          =   180
               Index           =   2
               Left            =   120
               TabIndex        =   20
               Top             =   75
               Width           =   720
            End
         End
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000003&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   2535
            Index           =   3
            Left            =   -74880
            ScaleHeight     =   2535
            ScaleWidth      =   4095
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   480
            Width           =   4095
            Begin VB.CheckBox chk 
               Appearance      =   0  'Flat
               BackColor       =   &H80000003&
               Caption         =   "全选"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   3
               Left            =   3240
               TabIndex        =   14
               Top             =   30
               Width           =   735
            End
            Begin MSComctlLib.ListView lvw 
               Height          =   1575
               Index           =   3
               Left            =   120
               TabIndex        =   15
               Top             =   480
               Width           =   3255
               _ExtentX        =   5741
               _ExtentY        =   2778
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   0   'False
               HideColumnHeaders=   -1  'True
               Checkboxes      =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   0
               NumItems        =   0
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "药品库房"
               Height          =   180
               Index           =   3
               Left            =   120
               TabIndex        =   16
               Top             =   75
               Width           =   720
            End
         End
         Begin VB.PictureBox pic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000003&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   2535
            Index           =   4
            Left            =   -74880
            ScaleHeight     =   2535
            ScaleWidth      =   4095
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   480
            Width           =   4095
            Begin VB.CheckBox chk 
               Appearance      =   0  'Flat
               BackColor       =   &H80000003&
               Caption         =   "全选"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   4
               Left            =   3240
               TabIndex        =   10
               Top             =   30
               Width           =   735
            End
            Begin MSComctlLib.ListView lvw 
               Height          =   1575
               Index           =   4
               Left            =   120
               TabIndex        =   11
               Top             =   480
               Width           =   3255
               _ExtentX        =   5741
               _ExtentY        =   2778
               View            =   3
               LabelEdit       =   1
               LabelWrap       =   -1  'True
               HideSelection   =   0   'False
               HideColumnHeaders=   -1  'True
               Checkboxes      =   -1  'True
               _Version        =   393217
               ForeColor       =   -2147483640
               BackColor       =   -2147483643
               BorderStyle     =   1
               Appearance      =   0
               NumItems        =   0
            End
            Begin VB.Label lbl 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "药品库房"
               Height          =   180
               Index           =   4
               Left            =   120
               TabIndex        =   12
               Top             =   75
               Width           =   720
            End
         End
      End
   End
   Begin VB.PictureBox picButton 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   5040
      ScaleHeight     =   615
      ScaleWidth      =   2895
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5520
      Width           =   2895
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   345
         Left            =   1320
         TabIndex        =   4
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "传送(&S)"
         Height          =   345
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.PictureBox picINF 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   120
      ScaleHeight     =   1095
      ScaleWidth      =   2055
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3480
      Width           =   2055
      Begin MSComctlLib.ListView lvwINF 
         Height          =   615
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   1085
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
   End
   Begin VB.PictureBox picClass 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   120
      ScaleHeight     =   2055
      ScaleWidth      =   2055
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   2055
      Begin MSComctlLib.ListView lvwClass 
         Height          =   1335
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   2355
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
   End
   Begin XtremeDockingPane.DockingPane dkpAreas 
      Left            =   120
      Top             =   120
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmTransmitBD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnShow As Boolean
Private mblnReturn As Boolean
Private mcolData As New Collection

Public Function ShowMe(ByVal frmOwner As Form, ByRef colData As Collection) As Boolean
    Set mcolData = Nothing
    Me.Show vbModal, frmOwner
    
    Set colData = mcolData
    ShowMe = mblnReturn
End Function

Private Sub chk_Click(Index As Integer)
    Dim l As Long
    
    If Me.Visible = False Then Exit Sub
    
    For l = 1 To lvw(Index).ListItems.Count
        lvw(Index).ListItems(l).Checked = chk(Index).Value = 1
    Next
    
End Sub

Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    Dim i As Byte
    Dim l As Long
    Dim blnFind As Boolean
    Dim strClass As String, strDetail As String, strINF As String
    Dim lsiTmp As ListItem
    
    '检查
    ''基础数据
    For i = 1 To lvwClass.ListItems.Count
        If lvwClass.ListItems(i).Checked Then
            blnFind = True
            Exit For
        End If
    Next
    If blnFind = False Then
        lvwClass.SetFocus
        MsgBox "请勾选“数据分类”的项目！", vbInformation, GSTR_MSG
        Exit Sub
    End If
    
    ''接口
    blnFind = False
    For i = 1 To lvwINF.ListItems.Count
        If lvwINF.ListItems(i).Checked Then
            blnFind = True
            Exit For
        End If
    Next
    If blnFind = False Then
        lvwINF.SetFocus
        MsgBox "请勾选“目标接口”！", vbInformation, GSTR_MSG
        Exit Sub
    End If
    
    
    
    '处理传送的数据
    '接口编号数据
    For i = 1 To lvwINF.ListItems.Count
        If lvwINF.ListItems(i).Checked Then
            strINF = strINF & lvwINF.ListItems(i).Tag & ";"
        End If
    Next
    If Right(strINF, 1) = ";" Then strINF = Left(strINF, Len(strINF) - 1)
    strINF = strINF & "|"
    
    blnFind = False
    For i = 1 To lvwClass.ListItems.Count
        If lvwClass.ListItems(i).Checked Then
            '数据分类
            strClass = i & "|"
            strDetail = ""
            For l = 1 To lvw(i - 1).ListItems.Count
                '数据明细
                Set lsiTmp = lvw(i - 1).ListItems(l)
                If lsiTmp.Checked Then
                    strDetail = strDetail & lsiTmp.Tag & ";"
                Else
                    If blnFind = False Then blnFind = True
                End If
            Next
            If Right(strDetail, 1) = ";" Then strDetail = Left(strDetail, Len(strDetail) - 1)
            
            '加入到集合对象
            ''格式：接口1[;接口n]|数据分类1[;数据分类n]|[[数据1][;数据n]]
            If blnFind Then
                mcolData.Add strINF & strClass & strDetail
            Else
                '全选默认strDetail空
                mcolData.Add strINF & strClass
            End If
        End If
    Next
    
    mblnReturn = True
    Me.Hide
End Sub

Private Sub dkpAreas_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case 1
        Item.Handle = picINF.hwnd
    Case 2
        Item.Handle = picButton.hwnd
    Case 3
        Item.Handle = picClass.hwnd
    Case 4
        Item.Handle = picSSB.hwnd
    End Select
End Sub

Private Sub Form_Activate()
    Dim i As Integer
    
    If mblnShow Then
        Screen.MousePointer = vbHourglass
        
        '加载数据
        Call FullData
        
        '调整SSTab
        Call sstClass_Click(0)
        
        lvwClass.SetFocus
        mblnShow = False
        
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub Form_Load()
    Dim i As Byte
    
    mblnReturn = False
    
    Call InitDockPane
        
    mblnShow = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then
        Cancel = True
        Me.Hide
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '
End Sub

Private Sub InitDockPane()
    Dim panBottomA As Pane, panBottomB As Pane, panClientLeft As Pane, panClient As Pane
    
    With dkpAreas
        .Options.UseSplitterTracker = False
        .Options.ThemedFloatingFrames = True
        .Options.AlphaDockingContext = True
        .Options.LunaColors = True
        .Options.HideClient = True
        .VisualTheme = ThemeOffice2003
        
        Set panBottomA = .CreatePane(1, 0, 50, DockBottomOf, panBottomA)
        With panBottomA
            .Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
            .Title = "目标接口"
            .MinTrackSize.Height = 50
        End With
        
        Set panBottomB = .CreatePane(2, 0, 40, DockBottomOf)
        With panBottomB
            .Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
            .Title = "按钮"
            .MaxTrackSize.Height = 40
            .MinTrackSize.Height = 40
        End With
        
        Set panClientLeft = .CreatePane(3, Me.ScaleY(Me.Height, vbTwips, vbPixels) \ 5 * 2, 150, DockTopOf)
        With panClientLeft
            .Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
            .Title = "数据分类"
            .MinTrackSize.Height = 100
            .MinTrackSize.Width = 100
        End With
        
        Set panClient = .CreatePane(4, 200, 0, DockRightOf, panClientLeft)
        With panClient
            .Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
            .Title = "右"
        End With
    End With
End Sub

Private Sub lvw_DblClick(Index As Integer)
    If Not lvw(Index).SelectedItem Is Nothing Then
        lvw(Index).SelectedItem.Checked = Not lvw(Index).SelectedItem.Checked
    End If
End Sub

Private Sub lvw_ItemCheck(Index As Integer, ByVal Item As MSComctlLib.ListItem)
    If Me.Visible Then
        If Not Item Is Nothing Then
            Item.Selected = True
        End If
    End If
End Sub

Private Sub lvwClass_Click()
    If Me.Visible = False Then Exit Sub
    
    If lvwClass.SelectedItem Is Nothing Then Exit Sub
    Call sstClass_Click(lvwClass.SelectedItem.Index - 1)
End Sub

Private Sub lvwClass_DblClick()
    If Not lvwClass.SelectedItem Is Nothing Then
        lvwClass.SelectedItem.Checked = Not lvwClass.SelectedItem.Checked
        Call sstClass_Click(lvwClass.SelectedItem.Index - 1)
    End If
End Sub

Private Sub lvwClass_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    If Me.Visible Then
        If Not Item Is Nothing Then
            Item.Selected = True
            Call lvwClass_Click
        End If
    End If
End Sub

'Private Sub lvwClass_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 32 Then
'        lvwClass_ItemCheck lvwClass.SelectedItem
'    End If
'End Sub

Private Sub pic_Resize(Index As Integer)
    On Error Resume Next
    With pic(Index)
        .Top = 0
        .Left = 0
        .Width = sstClass.Width
        .Height = sstClass.Height
    End With
    
    With lvw(Index)
        .Top = 330
        .Left = 0
        .Width = pic(Index).ScaleWidth
        .Height = pic(Index).ScaleHeight - lvw(Index).Top
    End With
    
    chk(Index).Left = pic(Index).ScaleWidth - chk(Index).Width - 120
End Sub

Private Sub picButton_Resize()
    On Error Resume Next
    With cmdCancel
        .Top = (picButton.ScaleHeight - .Height) \ 2
        .Left = picButton.ScaleWidth - .Width - 120
    End With
    With cmdOK
        .Top = cmdCancel.Top
        .Left = cmdCancel.Left - .Width - 120
    End With
End Sub

Private Sub picClass_Resize()
    On Error Resume Next
    With lvwClass
        .Top = 0
        .Left = 0
        .Width = picClass.ScaleWidth
        .Height = picClass.ScaleHeight
    End With
End Sub

Private Sub picINF_Resize()
    On Error Resume Next
    With lvwINF
        .Top = 0
        .Left = 0
        .Width = picINF.ScaleWidth
        .Height = picINF.ScaleHeight
    End With
End Sub

Private Sub picSSB_Resize()
    On Error Resume Next
    With sstClass
        .Top = 0
        .Left = 0
        .Width = picSSB.ScaleWidth
        .Height = picSSB.ScaleHeight
    End With
End Sub

Private Sub sstClass_Click(PreviousTab As Integer)
    Call pic_Resize(PreviousTab)
    
    If pic(PreviousTab).Tag <> "1" Then
        '未加载数据
        Call FullDataEx(PreviousTab)
    End If
    pic(PreviousTab).Enabled = lvwClass.SelectedItem.Checked
    pic(PreviousTab).ZOrder
End Sub

Private Sub FullData()
    Dim i As Byte
    Dim rsSQL As ADODB.Recordset
    Dim lsiTmp As ListItem

    '基础数据
    With lvwClass
        .ListItems.Clear
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , "DATA", "DATA", lvwClass.Width
        .SmallIcons = img16
        .Icons = img16
        .View = lvwReport
    End With
    
    For i = 0 To sstClass.Tabs - 1
        lvwClass.ListItems.Add , "K_" & i, sstClass.TabCaption(i), 1, 1
    Next
    
    '接口
    With lvwINF
        .ListItems.Clear
        .ColumnHeaders.Clear
        .SmallIcons = img24
        .Icons = img24
        .View = lvwIcon
    End With
    
    On Error GoTo hErr
    gstrSQL = _
            "Select ID, 编号, 名称 " & vbNewLine & _
            "From 药品设备接口 " & vbNewLine & _
            "Where 启用日期 Is Not Null And 停用日期 Is Null " & vbNewLine & _
            "Order By 编号 "
    Set rsSQL = gobjComLib.zlDatabase.OpenSQLRecord(gstrSQL, "获取启用的药品设备接口信息")
    Do While rsSQL.EOF = False
        Set lsiTmp = lvwINF.ListItems.Add(, "K_" & rsSQL!ID, rsSQL!编号, 1, 1)
        lsiTmp.ToolTipText = rsSQL!名称
        lsiTmp.Tag = CStr(rsSQL!编号)
        rsSQL.MoveNext
    Loop
    rsSQL.Close
    
    Exit Sub
    
hErr:
    If gobjComLib.ErrCenter = 1 Then Resume
End Sub

Private Sub FullDataEx(ByVal intIndex As Integer)
    Dim rsSQL As ADODB.Recordset
    Dim lsiTmp As ListItem
    Dim l As Long

    '初始化
    With lvw(intIndex)
        .View = lvwReport
        .ListItems.Clear
        .ColumnHeaders.Clear
        .ColumnHeaders.Add , "DATA", "DATA", lvw(intIndex).Width
        .FullRowSelect = True
    End With
    
    '获取数据
    Select Case intIndex
    Case Val("0-部门性质")
        gstrSQL = _
                "Select Distinct b.编码, b.名称" & vbNewLine & _
                "From 部门性质说明 A, 部门性质分类 B" & vbNewLine & _
                "Where a.工作性质 = b.名称" & vbNewLine & _
                "Order By b.名称"

    Case Val("1-人员工作性质")
        gstrSQL = _
                "Select 编码, 名称 From 人员性质分类 Order By 名称"
                
    Case Val("2-药品剂型")
        gstrSQL = _
                "Select 编码, 名称 From 药品剂型 Order By 名称"
                
    Case Val("3-药品库房"), Val("4-药品库房")
        gstrSQL = _
                "Select Distinct b.ID, b.名称" & vbNewLine & _
                "From 部门性质说明 A, 部门表 B" & vbNewLine & _
                "Where a.部门id = b.Id And a.工作性质 In ('西药库', '成药库', '中药库', '西药房', '成药房', '中药房')" & vbNewLine & _
                "Order By b.名称"
                
    End Select
    
    On Error GoTo hErr
    Set rsSQL = gobjComLib.zlDatabase.OpenSQLRecord(gstrSQL, "获取数据")
    
    '加载数据
    Do While rsSQL.EOF = False
        l = l + 1
        Set lsiTmp = lvw(intIndex).ListItems.Add(, "K_" & l, mdlMain.FormatString("[1]", rsSQL!名称))
        
        Select Case intIndex
        Case 0, 1
            lsiTmp.Tag = rsSQL!名称
        Case 2
            lsiTmp.Tag = rsSQL!名称
        Case 3, 4
            lsiTmp.Tag = rsSQL!ID
        End Select
        
        rsSQL.MoveNext
    Loop
    rsSQL.Close
    
    '完成并标记
    pic(intIndex).Tag = "1"
    
    Exit Sub
    
hErr:
    If gobjComLib.ErrCenter = 1 Then Resume
End Sub
