VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDefQueryItem 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "查询项目编辑"
   ClientHeight    =   3750
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7035
   Icon            =   "frmDefQueryItem.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   7035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin TabDlg.SSTab tbs 
      Height          =   3540
      Left            =   135
      TabIndex        =   33
      Top             =   90
      Width           =   5580
      _ExtentX        =   9843
      _ExtentY        =   6244
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      WordWrap        =   0   'False
      TabCaption(0)   =   "标题信息(&1)"
      TabPicture(0)   =   "frmDefQueryItem.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label6"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txt(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdOpen(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cbo(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txt(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cbo(1)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "文本(&2)"
      TabPicture(1)   =   "frmDefQueryItem.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "VisualTxt"
      Tab(1).Control(1)=   "cmdOpen(1)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "表格(&3)"
      TabPicture(2)   =   "frmDefQueryItem.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label4"
      Tab(2).Control(1)=   "Label9"
      Tab(2).Control(2)=   "msf"
      Tab(2).Control(3)=   "cboPos(0)"
      Tab(2).Control(4)=   "cbo(2)"
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "图形(&4)"
      TabPicture(3)   =   "frmDefQueryItem.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label5"
      Tab(3).Control(1)=   "lblSize(0)"
      Tab(3).Control(2)=   "UsrPicture(0)"
      Tab(3).Control(3)=   "cboPos(1)"
      Tab(3).Control(4)=   "cmdOpen(3)"
      Tab(3).Control(5)=   "cmdClear(0)"
      Tab(3).ControlCount=   6
      TabCaption(4)   =   "链接(&5)"
      TabPicture(4)   =   "frmDefQueryItem.frx":007C
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "cmdLeft"
      Tab(4).Control(1)=   "cmdRight"
      Tab(4).Control(2)=   "tvw"
      Tab(4).Control(3)=   "lvw"
      Tab(4).Control(4)=   "Label8"
      Tab(4).Control(5)=   "Label7"
      Tab(4).ControlCount=   6
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   2
         ItemData        =   "frmDefQueryItem.frx":0098
         Left            =   -74220
         List            =   "frmDefQueryItem.frx":009A
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   3120
         Width           =   2355
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "清除(&L)"
         Height          =   350
         Index           =   0
         Left            =   -70935
         TabIndex        =   21
         Top             =   900
         Width           =   1100
      End
      Begin VB.TextBox VisualTxt 
         Height          =   2670
         Left            =   -74895
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   390
         Width           =   5265
      End
      Begin VB.CommandButton cmdLeft 
         Caption         =   "<<"
         Height          =   350
         Left            =   -72240
         TabIndex        =   26
         Top             =   675
         Width           =   480
      End
      Begin VB.CommandButton cmdRight 
         Caption         =   ">>"
         Height          =   350
         Left            =   -72240
         TabIndex        =   27
         Top             =   1095
         Width           =   480
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "图片(&B)"
         Height          =   350
         Index           =   3
         Left            =   -70935
         TabIndex        =   20
         Top             =   435
         Width           =   1100
      End
      Begin VB.ComboBox cboPos 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   1
         ItemData        =   "frmDefQueryItem.frx":009C
         Left            =   -73740
         List            =   "frmDefQueryItem.frx":00A6
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   3090
         Width           =   2325
      End
      Begin VB.ComboBox cboPos 
         BackColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   0
         ItemData        =   "frmDefQueryItem.frx":00BA
         Left            =   -71145
         List            =   "frmDefQueryItem.frx":00C4
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   3120
         Width           =   1605
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "编辑(&E)"
         Height          =   350
         Index           =   1
         Left            =   -70695
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   3105
         Width           =   1100
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   1
         ItemData        =   "frmDefQueryItem.frx":00D8
         Left            =   1260
         List            =   "frmDefQueryItem.frx":00EE
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   3135
         Width           =   2760
      End
      Begin VB.Frame Frame2 
         Caption         =   "标题附加项"
         Height          =   1470
         Left            =   180
         TabIndex        =   34
         Top             =   1605
         Width           =   3855
         Begin VB.CommandButton cmdOpen 
            Caption         =   "…"
            Height          =   255
            Index           =   5
            Left            =   2175
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   990
            Width           =   270
         End
         Begin VB.CheckBox chk 
            Caption         =   "返回页首(&R)"
            Height          =   240
            Index           =   1
            Left            =   120
            TabIndex        =   8
            Top             =   660
            Width           =   1335
         End
         Begin VB.CheckBox chk 
            Caption         =   "隐藏标题(&H)"
            Height          =   240
            Index           =   0
            Left            =   135
            TabIndex        =   7
            Top             =   315
            Width           =   1335
         End
         Begin VB.CheckBox chk 
            Caption         =   "标题图标(&I)"
            Height          =   240
            Index           =   2
            Left            =   105
            TabIndex        =   9
            Top             =   1005
            Width           =   1335
         End
         Begin zl9NewQuery.ctlPicture UsrPicture 
            Height          =   435
            Index           =   1
            Left            =   1455
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   900
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   767
         End
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   0
         Left            =   1245
         MaxLength       =   30
         TabIndex        =   1
         Top             =   450
         Width           =   2760
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   0
         Left            =   1245
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1230
         Width           =   2760
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "…"
         Height          =   255
         Index           =   0
         Left            =   3720
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   885
         Width           =   270
      End
      Begin zl9NewQuery.ctlPicture UsrPicture 
         Height          =   2505
         Index           =   0
         Left            =   -74820
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   435
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   4419
      End
      Begin MSComctlLib.TreeView tvw 
         Height          =   2700
         Left            =   -71715
         TabIndex        =   29
         Top             =   675
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   4763
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   494
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         ImageList       =   "ils16"
         Appearance      =   1
      End
      Begin MSComctlLib.ListView lvw 
         Height          =   2670
         Left            =   -74910
         TabIndex        =   25
         Top             =   675
         Width           =   2610
         _ExtentX        =   4604
         _ExtentY        =   4710
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ils16"
         SmallIcons      =   "ils16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "链接页面"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "页面项目"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.TextBox txt 
         BackColor       =   &H00EEEEEE&
         Height          =   300
         Index           =   1
         Left            =   1245
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   855
         Width           =   2760
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msf 
         Height          =   2565
         Left            =   -74895
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   435
         Width           =   5370
         _ExtentX        =   9472
         _ExtentY        =   4524
         _Version        =   393216
         FixedRows       =   0
         FixedCols       =   0
         BackColorBkg    =   -2147483628
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         MergeCells      =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "表格(&T)"
         Height          =   180
         Left            =   -74895
         TabIndex        =   16
         Top             =   3165
         Width           =   630
      End
      Begin VB.Label Label8 
         Caption         =   "连接项目(&I)"
         Height          =   210
         Left            =   -74895
         TabIndex        =   24
         Top             =   465
         Width           =   1500
      End
      Begin VB.Label Label7 
         Caption         =   "供选择的页面(&P)"
         Height          =   240
         Left            =   -71700
         TabIndex        =   28
         Top             =   465
         Width           =   1620
      End
      Begin VB.Label lblSize 
         AutoSize        =   -1  'True
         Caption         =   "800 X 600"
         Height          =   180
         Index           =   0
         Left            =   -70995
         TabIndex        =   36
         Top             =   2520
         Width           =   810
      End
      Begin VB.Label Label5 
         Caption         =   "插图位置(&Y)"
         Height          =   195
         Left            =   -74835
         TabIndex        =   22
         Top             =   3150
         Width           =   1545
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "位置(&X)"
         Height          =   180
         Left            =   -71760
         TabIndex        =   18
         Top             =   3180
         Width           =   630
      End
      Begin VB.Label Label6 
         Caption         =   "项目类型(&T)"
         Height          =   195
         Left            =   180
         TabIndex        =   11
         Top             =   3195
         Width           =   1545
      End
      Begin VB.Label Label1 
         Caption         =   "标题文本(&L)"
         Height          =   225
         Left            =   165
         TabIndex        =   0
         Top             =   510
         Width           =   1680
      End
      Begin VB.Label Label2 
         Caption         =   "标题字体(&F)"
         Height          =   210
         Left            =   165
         TabIndex        =   2
         Top             =   900
         Width           =   1725
      End
      Begin VB.Label Label3 
         Caption         =   "标题位置(&A)"
         Height          =   195
         Left            =   165
         TabIndex        =   5
         Top             =   1275
         Width           =   1545
      End
   End
   Begin MSComDlg.CommonDialog dlg 
      Left            =   7440
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   5850
      TabIndex        =   30
      Top             =   405
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   5850
      TabIndex        =   31
      Top             =   825
      Width           =   1100
   End
   Begin VB.CommandButton Command1 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   5850
      TabIndex        =   32
      Top             =   1470
      Width           =   1100
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   4605
      Top             =   3915
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
            Picture         =   "frmDefQueryItem.frx":0146
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDefQueryItem.frx":04E0
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmDefQueryItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnFirst As Boolean
Private mKey As Long
Private mOrder As Long
Private mOK As Boolean

Private mvarSvrPicRange As String           '保存增加图片的范围
Private mvarSvrPicType As String            '保存增加图片的类型

Private objTxt As TextBox

Private Sub cbo_Click(Index As Integer)
        
    Select Case Index
    Case 1
        tbs.TabEnabled(1) = False
        tbs.TabEnabled(2) = False
        tbs.TabEnabled(3) = False
        tbs.TabEnabled(4) = False
        
        Select Case cbo(Index).ItemData(cbo(Index).ListIndex)
        Case 0
            tbs.TabEnabled(1) = True
            Call DisableObject
        Case 1
            tbs.TabEnabled(2) = True
            Call DisableObject
            cboPos(0).ListIndex = 0
            cboPos(0).Enabled = False
        Case 2
            tbs.TabEnabled(3) = True
            Call DisableObject
            cboPos(1).ListIndex = 0
            cboPos(1).Enabled = False
        Case 3
            tbs.TabEnabled(4) = True
            Call DisableObject
        Case 4
            tbs.TabEnabled(1) = True
            tbs.TabEnabled(2) = True
            Call DisableObject
            cboPos(0).Enabled = True
        Case 5
            tbs.TabEnabled(1) = True
            tbs.TabEnabled(3) = True
            Call DisableObject
            cboPos(1).Enabled = True
        End Select
    Case 2
        If mblnFirst = False Then Call ShowTable(cbo(Index).ItemData(cbo(Index).ListIndex))
    End Select
    cmdOK.Tag = "1"
End Sub

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 23 Then
        KeyAscii = 0
        
        Select Case Index
        Case 1
            If tbs.TabEnabled(1) Then
                tbs.Tab = 1
                VisualTxt.SetFocus
            ElseIf tbs.TabEnabled(2) Then
                tbs.Tab = 2
                cbo(2).SetFocus
            ElseIf tbs.TabEnabled(3) Then
                tbs.Tab = 3
                cmdOpen(3).SetFocus
            ElseIf tbs.TabEnabled(4) Then
                tbs.Tab = 4
                lvw.SetFocus
            Else
                cmdOK.SetFocus
            End If
        Case 2
            If cboPos(0).Enabled Then
                SendKeys "{TAB}"
            Else
                cmdOK.SetFocus
            End If
        Case Else
            
            SendKeys "{TAB}"
            
        End Select
        
        
    End If
    
End Sub

Private Sub cboPos_Click(Index As Integer)
    cmdOK.Tag = "1"
End Sub

Private Sub cboPos_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Select Case Index
        Case 0
            If tbs.TabEnabled(3) Then
                tbs.Tab = 3
                cmdOpen(3).SetFocus
            ElseIf tbs.TabEnabled(4) Then
                tbs.Tab = 4
                lvw.SetFocus
            Else
                cmdOK.SetFocus
            End If
        Case 1
            If tbs.TabEnabled(4) Then
                tbs.Tab = 4
                lvw.SetFocus
            Else
                cmdOK.SetFocus
            End If
        End Select
    End If

End Sub

Private Sub chk_Click(Index As Integer)
    cmdOK.Tag = "1"
    Select Case Index
    Case 2
        UsrPicture(1).Tag = ""
        UsrPicture(1).Cls
        cmdOpen(5).Visible = chk(Index).Value
    End Select
End Sub

Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
        Exit Sub
    End If
    If Chr(KeyAscii) = "*" Then
        KeyAscii = 0
        If Index = 2 Then
            If cmdOpen(5).Enabled Then Call cmdOpen_Click(5)
        End If
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdClear_Click(Index As Integer)
    UsrPicture(0).Tag = ""
    UsrPicture(0).Cls
End Sub

Private Sub cmdClear_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        Select Case Index
        Case 0
            cmdOK.SetFocus
        End Select
    End If
End Sub

Private Sub cmdLeft_Click()
    Dim PageNo As Long
    Dim Order As Long
    Dim Itmx As ListItem
    
    
    If tvw.SelectedItem Is Nothing Then Exit Sub
    'If tvw.SelectedItem.Image <> 2 Then Exit Sub
        
    If tvw.SelectedItem.Image <> 2 Then
        PageNo = Val(Mid(tvw.SelectedItem.Key, 2))
        Order = 0
        
        '检查当前项目是否已经添加，如果已经添加，则不能重复添加
        If CheckIn(tvw.SelectedItem.Key & "C0") = True Then Exit Sub
        Set Itmx = lvw.ListItems.Add(, "K" & PageNo & "C" & Order, tvw.SelectedItem.Text, 2, 2)
        Itmx.SubItems(1) = ""
    Else
        PageNo = Val(Mid(tvw.SelectedItem.Key, 2, InStr(tvw.SelectedItem.Key, "C") - 2))
        Order = Val(Mid(tvw.SelectedItem.Key, InStr(tvw.SelectedItem.Key, "C") + 1))
        
        '检查当前项目是否已经添加，如果已经添加，则不能重复添加
        If CheckIn(tvw.SelectedItem.Key) = True Then Exit Sub
        Set Itmx = lvw.ListItems.Add(, "K" & PageNo & "C" & Order, tvw.SelectedItem.Parent.Text, 2, 2)
        Itmx.SubItems(1) = tvw.SelectedItem.Text
    End If
    
    cmdOK.Tag = "1"
End Sub

Private Sub cmdOK_Click()
    If SaveData Then
        mOK = True
        If mOrder = 0 Then
            Call RestoreEditState
            cmdOK.Tag = ""
            txt(0).SetFocus
        Else
            cmdOK.Tag = ""
            Unload Me
        End If
    End If
End Sub

Private Sub cmdOpen_Click(Index As Integer)
    Dim lngKey As Long
    
    Select Case Index
    Case 0
        On Error Resume Next
        dlg.CancelError = True
        dlg.flags = &H3 Or &H100 Or &H400 Or &H200 Or &H10000
        
        dlg.FontName = txt(1).FontName
        dlg.FontSize = txt(1).FontSize
        dlg.FontBold = txt(1).FontBold
        dlg.FontItalic = txt(1).FontItalic
        dlg.Color = txt(1).ForeColor
        dlg.ShowFont
        If Err.Number = 0 Then
            txt(1).FontName = dlg.FontName
            txt(1).FontSize = dlg.FontSize
            txt(1).FontBold = dlg.FontBold
            txt(1).FontItalic = dlg.FontItalic
            txt(1).ForeColor = dlg.Color
            txt(1).Text = txt(1).FontName & ";" & txt(1).FontSize & ";" & IIf(txt(1).FontBold, "1", "0") & ";" & IIf(txt(1).FontItalic, "1", "0") & ";" & txt(1).ForeColor
        Else
            Err.Clear
        End If
        On Error GoTo 0
    Case 1
        If frmTextEdit.OpenTextEditDialog(Me, VisualTxt) Then
            cmdOK.Tag = "1"
        End If
    Case 3
        If frmPicSelect.OpenPictureBox(Me, "添加图片", "9;0;1;2;3;4", lngKey, mvarSvrPicRange, mvarSvrPicType) Then
            '更新图片显示
            UsrPicture(0).Tag = lngKey
            Call ShowPicture(lngKey, 0)
            cmdOK.Tag = "1"
            SendKeys "{TAB}"
        End If
    Case 5
        If frmPicSelect.OpenPictureBox(Me, "添加项目图标", "3", lngKey, mvarSvrPicRange, mvarSvrPicType) Then
            '更新图片显示
            UsrPicture(1).Tag = lngKey
            Call ShowPicture(lngKey, 1)
            cmdOK.Tag = "1"
        End If
    End Select
End Sub


Private Sub cmdRight_Click()
    If lvw.SelectedItem Is Nothing Then Exit Sub
    
    lvw.ListItems.Remove lvw.SelectedItem.Index
    
    cmdOK.Tag = "1"
End Sub

Private Sub Command1_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub Form_Activate()
    Dim i As Integer
    Dim strTmp As String
    Dim W As Single
    Dim H As Single
    Dim vTmp As String
    
    If mblnFirst = False Then Exit Sub
    DoEvents
    
    '初始化过程
    cbo(0).Clear
    cbo(0).AddItem "0-左对齐"
    cbo(0).AddItem "1-右对齐"
    cbo(0).AddItem "2-居中"
    cbo(0).ListIndex = 0
    cbo(1).ListIndex = 0
    
    Call LoadTable
    
    cboPos(0).ListIndex = 0
    cboPos(1).ListIndex = 0
    
    
    mblnFirst = False
    
    On Error GoTo errHand
           
    Call RestoreEditState
    
    If mOrder > 0 Then
        gstrSQL = "select 段落类型,段落字体,标题文本,标题位置,标题字体,标题隐藏,返回页首,标题图标,插表序号,插图序号,插表位置,插图位置 from 咨询段落目录 A where A.页面序号=[1] and A.段落序号=[2]"
        Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mKey, mOrder)
        If gRs.BOF = False Then
            txt(0).Text = IIf(IsNull(gRs!标题文本), "", gRs!标题文本)
            cbo(0).ListIndex = IIf(IsNull(gRs!标题位置), 0, gRs!标题位置)
            
            txt(1).Text = IIf(IsNull(gRs!标题字体), "宋体;12;0;0;0", gRs!标题字体)
            txt(1).FontName = Split(txt(1).Text, ";")(0)
            txt(1).FontSize = Val(Split(txt(1).Text, ";")(1))
            txt(1).FontBold = Val(Split(txt(1).Text, ";")(2))
            txt(1).FontItalic = Val(Split(txt(1).Text, ";")(3))
            txt(1).ForeColor = Val(Split(txt(1).Text, ";")(4))
            
            chk(0).Value = IIf(IsNull(gRs!标题隐藏), 0, gRs!标题隐藏)
            
            chk(1).Value = IIf(IsNull(gRs!返回页首), 0, gRs!返回页首)
            
            If IsNull(gRs!标题图标) = False Then
                chk(2).Value = 1
                UsrPicture(1).Tag = gRs!标题图标
                Call ShowPicture(gRs!标题图标, 1)
            Else
                chk(2).Value = 0
            End If
            
            tbs.TabEnabled(1) = False
            
            VisualTxt.Text = Sys.ReadLob(glngSys, 29, mKey & "," & mOrder, "", 1)
            
            If VisualTxt.Text <> "" Then
                tbs.TabEnabled(1) = True
                strTmp = IIf(IsNull(gRs!段落字体), "宋体;12;0;0;0", gRs!段落字体)
'                VisualTxt.Text = gRs!段落文本
                VisualTxt.FontName = Split(strTmp, ";")(0)
                VisualTxt.FontSize = Split(strTmp, ";")(1)
                VisualTxt.FontBold = Split(strTmp, ";")(2)
                VisualTxt.FontItalic = Split(strTmp, ";")(3)
                VisualTxt.ForeColor = Split(strTmp, ";")(4)
            End If
            
            If IsNull(gRs!插表序号) = False Then
                tbs.TabEnabled(2) = True
                cboPos(0).ListIndex = IIf(IsNull(gRs!插表位置), 0, gRs!插表位置)
                cbo(2).ListIndex = FindCboIndex(cbo(2), gRs!插表序号)
                'Call ShowTable(gRs!插表序号)
            End If
            
            If IsNull(gRs!插图序号) = False Then
                tbs.TabEnabled(3) = True
                cboPos(1).ListIndex = IIf(IsNull(gRs!插图位置), 0, gRs!插图位置)
                UsrPicture(0).Tag = gRs!插图序号
                ShowPicture Val(UsrPicture(0).Tag), 0
            End If
            tbs.TabEnabled(4) = IIf(LoadFirstSuperConnect(mKey, mOrder), 1, 0)
                        
            If tbs.TabEnabled(1) = True And tbs.TabEnabled(2) = False And tbs.TabEnabled(3) = False And tbs.TabEnabled(4) = False Then cbo(1).ListIndex = 0
            If tbs.TabEnabled(1) = False And tbs.TabEnabled(2) = True And tbs.TabEnabled(3) = False And tbs.TabEnabled(4) = False Then cbo(1).ListIndex = 1
            If tbs.TabEnabled(1) = False And tbs.TabEnabled(2) = False And tbs.TabEnabled(3) = True And tbs.TabEnabled(4) = False Then cbo(1).ListIndex = 2
            If tbs.TabEnabled(1) = False And tbs.TabEnabled(2) = False And tbs.TabEnabled(3) = False And tbs.TabEnabled(4) = True Then cbo(1).ListIndex = 3
            If tbs.TabEnabled(1) = True And tbs.TabEnabled(2) = True And tbs.TabEnabled(3) = False And tbs.TabEnabled(4) = False Then cbo(1).ListIndex = 4
            If tbs.TabEnabled(1) = True And tbs.TabEnabled(2) = False And tbs.TabEnabled(3) = True And tbs.TabEnabled(4) = False Then cbo(1).ListIndex = 5
            If tbs.TabEnabled(1) = False And tbs.TabEnabled(2) = False And tbs.TabEnabled(3) = False And tbs.TabEnabled(4) = False Then
               cbo(1).ListIndex = 0
               tbs.TabEnabled(1) = True
            End If
            
        End If
    End If
    
    Call LoadPageTree
    If tvw.Nodes.Count > 0 Then tvw.Nodes(1).Expanded = True
    
    cmdOK.Tag = ""
    
    If mOrder > 0 And frmDefQuery.lvw.SelectedItem.Tag = "1" Then
        '为固定查询页面的查询项,不能编辑标题信息
        txt(0).Enabled = False
        txt(1).Enabled = False
        cmdOpen(0).Enabled = False
        cbo(0).Enabled = False
        chk(0).Enabled = False
        chk(1).Enabled = False
        chk(2).Enabled = False
        cbo(1).Enabled = False
        cmdOpen(5).Enabled = False
        
    End If
    
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    cmdOK.Tag = ""
End Sub

Private Sub Form_Load()
    mblnFirst = True
    mOK = False
    
    '初始化数据或控件属性
    mvarSvrPicRange = ""
    mvarSvrPicType = ""
End Sub

Public Function ShowItemEdit(frmMain As Object, ByVal Key As Long, ByVal Order As Long) As Boolean
    mKey = Key
    mOrder = Order
    frmDefQueryItem.Show 1, frmMain
    ShowItemEdit = mOK
End Function

Private Function SaveData() As Boolean
    Dim lng段号 As Long
    Dim i As Long
    Dim strTable As String
    Dim strPic As String
    Dim strSQL() As String
    Dim PageNo As Long
    Dim OrderNo As Long
    Dim strFont As String
    Dim rs As New ADODB.Recordset

    ReDim strSQL(1 To 2 + lvw.ListItems.Count)

    If cmdOK.Tag <> "" Then
                
        strTable = IIf(tbs.TabEnabled(2) = True, msf.Tag & ";" & cboPos(0).ListIndex, ";")
        strPic = IIf(tbs.TabEnabled(3) = True, UsrPicture(0).Tag & ";" & cboPos(1).ListIndex, ";")
        
        If tbs.TabEnabled(1) = True Then strFont = VisualTxt.FontName & ";" & VisualTxt.FontSize & ";" & VisualTxt.FontBold & ";" & VisualTxt.FontItalic & ";" & VisualTxt.ForeColor
        If mOrder = 0 Then
            lng段号 = CalcOrder(mKey)
            strSQL(1) = "zl_咨询段落目录_insert(" & mKey & "," & lng段号 & ",'" & txt(0).Text & "'," & IIf(chk(2).Value = 1, IIf(Val(UsrPicture(1).Tag) = 0, "NULL", Val(UsrPicture(1).Tag)), 0) & "," & chk(0).Value & "," & cbo(0).ListIndex & ",'" & IIf(txt(1).Text = "", "宋体;12;0;0;0", txt(1).Text) & "'," & chk(1).Value & "," & Val(Split(strTable, ";")(0)) & "," & Val(Split(strTable, ";")(1)) & "," & Val(Split(strPic, ";")(0)) & "," & Val(Split(strPic, ";")(1)) & ",'" & strFont & "'," & Left(cbo(1).Text, 1) & ")"
        Else
            lng段号 = mOrder
            strSQL(1) = "zl_咨询段落目录_update(" & mKey & "," & lng段号 & ",'" & txt(0).Text & "'," & IIf(chk(2).Value = 1, IIf(Val(UsrPicture(1).Tag) = 0, "NULL", Val(UsrPicture(1).Tag)), 0) & "," & chk(0).Value & "," & cbo(0).ListIndex & ",'" & txt(1).Text & "'," & chk(1).Value & "," & Val(Split(strTable, ";")(0)) & "," & Val(Split(strTable, ";")(1)) & "," & Val(Split(strPic, ";")(0)) & "," & Val(Split(strPic, ";")(1)) & ",'" & strFont & "'," & Left(cbo(1).Text, 1) & ")"
            strSQL(2) = "zl_咨询段落链接_delete(" & mKey & "," & lng段号 & ")"
        End If
        If tbs.TabEnabled(4) = True Then
            For i = 1 To lvw.ListItems.Count
                PageNo = Val(Mid(lvw.ListItems(i).Key, 2, InStr(lvw.ListItems(i).Key, "C") - 2))
                OrderNo = Val(Mid(lvw.ListItems(i).Key, InStr(lvw.ListItems(i).Key, "C") + 1))
                strSQL(2 + i) = "zl_咨询段落链接_insert(" & mKey & "," & lng段号 & "," & PageNo & "," & OrderNo & ")"
            Next
        End If

        On Error GoTo errHand
        gcnOracle.BeginTrans
        For i = 1 To 2 + lvw.ListItems.Count
            If strSQL(i) <> "" Then Call zlDatabase.ExecuteProcedure(strSQL(i), Me.Caption)
        Next

        If tbs.TabEnabled(1) = True Then
            '保存大文本内容
            Call Sys.SaveLob(glngSys, 29, mKey & "," & lng段号, VisualTxt.Text, 1)
        End If
        gcnOracle.CommitTrans
        Call frmDefQuery.RefreshItem(CStr(lng段号))
    End If
    
    SaveData = True
    Exit Function
errHand:
    
    gcnOracle.RollbackTrans
    If ErrCenter() = -1 Then Resume
    
End Function

Private Sub RestoreEditState()
    Dim i As Long
    
    For i = 0 To txt.UBound
        txt(i).Text = ""
        txt(i).Tag = ""
    Next
    
    chk(0).Value = 0
    chk(1).Value = 0
    chk(2).Value = 1
        
    txt(1).Text = "宋体;12;0;0;0"
    txt(1).FontSize = 12
        
    VisualTxt.Text = ""
    UsrPicture(0).Tag = ""
    UsrPicture(0).Cls
    msf.Rows = 1
    ClearSpecRowCol msf, 0, Array()
    lvw.ListItems.Clear
    
End Sub

Private Function LoadFirstSuperConnect(ByVal PageNo As Long, ByVal Order As Long) As Boolean
'加载第一个超级连接项目名称
    Dim Itmx As ListItem
    
    gstrSQL = "select A.链接页面,A.页内段号,B.页面名称 from 咨询段落链接 A,咨询页面目录 B where A.链接页面=B.页面序号 AND  A.页面序号=[1] and A.段落序号=[2]"
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, PageNo, Order)
    If gRs.BOF = False Then
        While Not gRs.EOF
            Set Itmx = lvw.ListItems.Add(, "K" & gRs!链接页面 & "C" & gRs!页内段号, IIf(IsNull(gRs!页面名称), "", gRs!页面名称), 2, 2)
            Itmx.SubItems(1) = LoadPageItemName(gRs!链接页面, gRs!页内段号)
            gRs.MoveNext
        Wend
    End If
    LoadFirstSuperConnect = IIf(lvw.ListItems.Count > 0, True, False)
End Function

Private Function CalcOrder(ByVal PageNo As Long) As Long
'计算页内的序号值
    CalcOrder = 0
    gstrSQL = "select nvl(max(段落序号),0)+1 from 咨询段落目录 where 页面序号=[1]"
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, PageNo)
    If gRs.BOF = False Then CalcOrder = gRs.Fields(0).Value
End Function

Private Sub Form_Unload(Cancel As Integer)
    If cmdOK.Tag = "1" Then
        If MsgBox("查询项目已经更改，确认不保存就退出吗？", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then Cancel = True
    End If
End Sub



Private Sub txt_Change(Index As Integer)
    cmdOK.Tag = "1"
End Sub

Private Function LoadPageItemName(ByVal PageNo As Long, ByVal Order As Long) As String
    Dim rs As New ADODB.Recordset
    
    gstrSQL = "select 页面序号 from 咨询页面目录 where 页面名称='专家介绍'"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If rs.BOF = False Then
        If rs!页面序号 = PageNo Then
            gstrSQL = "select A.姓名||'['||C.名称||']' as result from 人员表 A,部门人员 B,部门表 C where A.id=B.人员id And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) and B.部门id=C.id and A.id=[1]"
        Else
            gstrSQL = "select 标题文本 as result from 咨询段落目录 where 页面序号=[2] and 段落序号=[1]"
        End If
    Else
        gstrSQL = "select 标题文本 as result from 咨询段落目录 where 页面序号=[2] and 段落序号=[1]"
    End If
    
    If gstrSQL <> "" Then
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Order, PageNo)
        If rs.BOF = False Then
            LoadPageItemName = IIf(IsNull(rs!Result), "", rs!Result)
        End If
    End If
End Function

Private Sub txt_GotFocus(Index As Integer)
    If Index = 0 Then zlCommFun.OpenIme True
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
        Exit Sub
    End If
    
    If Index = 0 Then
        If CheckIsInclude(UCase(Chr(KeyAscii)), "'") = True Then KeyAscii = 0
    End If
    
    If Chr(KeyAscii) = "*" Then
        KeyAscii = 0
        Select Case Index
        Case 1
            If cmdOpen(0).Enabled Then Call cmdOpen_Click(0)
        Case 2
            If cmdOpen(1).Enabled Then Call cmdOpen_Click(1)
        Case 3
            If cmdOpen(2).Enabled Then Call cmdOpen_Click(2)
        Case 4
            If cmdOpen(3).Enabled Then Call cmdOpen_Click(3)
        Case 5
            If cmdOpen(4).Enabled Then Call cmdOpen_Click(4)
        End Select
    End If
    
End Sub

Private Sub txt_LostFocus(Index As Integer)
    If Index = 0 Then zlCommFun.OpenIme
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    Cancel = Not StrIsValid(txt(Index).Text, txt(Index).MaxLength)
End Sub

Private Sub LoadPageTree()
    '加载页面数据及页面的组成项目
    Dim nodx As Node
    
    gstrSQL = "select 页面序号,页面名称,固定页面 from 咨询页面目录 where 页面序号>0 and 页面序号<>[1]"
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mKey)
    If gRs.BOF = False Then
        While Not gRs.EOF
            Set nodx = tvw.Nodes.Add(, , "K" & gRs!页面序号, IIf(IsNull(gRs!页面名称), "", gRs!页面名称), 1, 1)
            nodx.Tag = IIf(IsNull(gRs!固定页面), 0, gRs!固定页面)
            If nodx.Text = "专家介绍" Then Call LoadPersonList(Val(Mid(nodx.Key, 2)))
            Call LoadPageItem(Val(Mid(nodx.Key, 2)))
            gRs.MoveNext
        Wend
    End If
End Sub

Private Sub LoadPageItem(ByVal PageNo As Long)
    Dim rs As New ADODB.Recordset
    Dim nodx As Node
    
    gstrSQL = "select 段落序号,标题文本 from 咨询段落目录 where 页面序号=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, PageNo)
    If rs.BOF = False Then
        While Not rs.EOF
            Set nodx = tvw.Nodes.Add("K" & PageNo, tvwChild, "K" & PageNo & "C" & rs!段落序号, IIf(IsNull(rs!标题文本), "", rs!标题文本), 2, 2)
            rs.MoveNext
        Wend
    End If
    CloseRecord rs
End Sub

Private Sub LoadPersonList(ByVal PageNo As Long)
    Dim nodx As Node
    Dim rs As New ADODB.Recordset
    
    gstrSQL = "select D.人员id,A.姓名,B.名称 as 部门 from 人员表 A,部门表 B,部门人员 C,咨询专家清单 D where D.人员id=C.人员id And (a.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.撤档时间 Is Null) and D.科室id=C.部门id and C.缺省=1 and A.id=C.人员id and B.id=C.部门id"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If rs.BOF = False Then
        While Not rs.EOF
            Set nodx = tvw.Nodes.Add("K" & PageNo, tvwChild, "K" & PageNo & "C" & rs!人员ID, IIf(IsNull(rs!姓名), "", IIf(IsNull(rs!部门), "", IIf(IsNull(rs!姓名), "", rs!姓名) & "[" & rs!部门 & "]")), 2, 2)
            rs.MoveNext
        Wend
    End If
    CloseRecord rs
End Sub

Private Function CheckIn(ByVal Key As String) As Boolean
    Dim i As Long
    
    For i = 1 To lvw.ListItems.Count
        If lvw.ListItems(i).Key = Key Then
            CheckIn = True
            Exit Function
        End If
    Next
    CheckIn = False
End Function

Private Sub ShowTable(ByVal No As Long)
    '显示表格到界面上
    Dim i As Long
    Dim strTmp As String
    Dim intPos As Long
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
    msf.Tag = 0
    
    gstrSQL = "select 序号,名称,列数,列宽,行数,行高,合并行,合并列,颜色 from 咨询表格元素 where 序号=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, No)
    If rs.BOF = False Then
        msf.Tag = No
        If IIf(IsNull(rs!行数), 0, rs!行数) <= 0 Then
            MsgBox "错误的表格行数（行数小于1）！", vbInformation, gstrSysName
            Exit Sub
        End If
        If IIf(IsNull(rs!列数), 0, rs!列数) <= 0 Then
            MsgBox "错误的表格列数（行数小于1）！", vbInformation, gstrSysName
            Exit Sub
        End If
                
        msf.Rows = rs!行数
        msf.Cols = rs!列数
        
        On Error Resume Next
        For i = 0 To msf.Rows - 1
            msf.RowHeight(i) = 300
        Next
        For i = 0 To msf.Rows - 1
            msf.RowHeight(i) = Split(rs!行高, ";")(i)
        Next
                        
        For i = 0 To msf.Cols - 1
            msf.ColWidth(i) = 1200
        Next
        For i = 0 To msf.Cols - 1
            msf.ColWidth(i) = Split(rs!列宽, ";")(i)
        Next
                                
        strTmp = IIf(IsNull(rs!合并行), "", rs!合并行 & ";")
        intPos = InStr(strTmp, ";")
        While intPos > 0
            msf.MergeRow(Val(Mid(strTmp, 1, intPos - 1)) - 1) = True
            strTmp = Mid(strTmp, intPos + 1)
            intPos = InStr(strTmp, ";")
        Wend

        strTmp = IIf(IsNull(rs!合并列), "", rs!合并列 & ";")
        intPos = InStr(strTmp, ";")
        While intPos > 0
            msf.MergeCol(Val(Mid(strTmp, 1, intPos - 1)) - 1) = True
            strTmp = Mid(strTmp, intPos + 1)
            intPos = InStr(strTmp, ";")
        Wend
        
        gstrSQL = "select 表号,行号,列号,内容,对齐,颜色,字体 from 咨询表格内容 where 表号=[1]"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, No)
        If rs.BOF = False Then
            While Not rs.EOF
                msf.Row = rs!行号 - 1
                msf.Col = rs!列号 - 1
                msf.TextMatrix(msf.Row, msf.Col) = IIf(IsNull(rs!内容), "", rs!内容)
                msf.CellAlignment = IIf(IsNull(rs!对齐), 9, rs!对齐)
                msf.CellForeColor = IIf(IsNull(rs!颜色), 0, rs!颜色)
                msf.CellFontName = Split(IIf(IsNull(rs!字体), "宋体;9;False;False;False;False", rs!字体), ";")(0)
                msf.CellFontSize = Split(IIf(IsNull(rs!字体), "宋体;9;False;False;False;False", rs!字体), ";")(1)
                msf.CellFontBold = IIf(Split(IIf(IsNull(rs!字体), "宋体;9;False;False;False;False", rs!字体), ";")(2) = True, True, False)
                msf.CellFontItalic = IIf(Split(IIf(IsNull(rs!字体), "宋体;9;False;False;False;False", rs!字体), ";")(3) = True, True, False)
                msf.CellFontStrikeThrough = IIf(Split(IIf(IsNull(rs!字体), "宋体;9;False;False;False;False", rs!字体), ";")(4) = True, True, False)
                msf.CellFontUnderline = IIf(Split(IIf(IsNull(rs!字体), "宋体;9;False;False;False;False", rs!字体), ";")(5) = True, True, False)
                rs.MoveNext
            Wend
        End If
        msf.Visible = True
    End If
    CloseRecord rs
    
    Exit Sub
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadTable()
    
    gstrSQL = "select 序号,名称 from 咨询表格元素"
    
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    If gRs.BOF = False Then
        While Not gRs.EOF
            cbo(2).AddItem IIf(IsNull(gRs!名称), "", gRs!名称)
            cbo(2).ItemData(cbo(2).NewIndex) = gRs!序号
            gRs.MoveNext
        Wend
    End If
End Sub

Private Sub ShowPicture(ByVal PicNo As Long, ByVal Index As Long)
    Dim rs As New ADODB.Recordset
    
    gstrSQL = "select 序号,宽度,高度,类型 from 咨询图片元素 where 序号=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, PicNo)
    If rs.BOF = False Then
        Call UsrPicture(Index).ShowPictureByFieldNew(rs!序号, rs!宽度 * Screen.TwipsPerPixelX, rs!高度 * Screen.TwipsPerPixelY, IIf(IsNull(rs!类型), 0, rs!类型))
        If Index = 0 Then lblSize(Index).Caption = "宽度:" & Format(rs!宽度 * Screen.TwipsPerPixelX / 567, "0.0(厘米)") & vbCrLf & "高度:" & Format(rs!高度 * Screen.TwipsPerPixelY / 567, "0.0(厘米)")
    End If
    CloseRecord rs
End Sub

Private Function CheckItemLimit(ByVal PageNo As Long) As Boolean
    gstrSQL = "select nvl(count(*),0) from 咨询段落目录 where 页面序号=[1]"
    Set gRs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, PageNo)
    If gRs.BOF = False Then CheckItemLimit = IIf(gRs.Fields(0).Value < 16, True, False)
End Function

Private Sub VisualTxt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If tbs.TabEnabled(2) Then
            tbs.Tab = 2
            cbo(2).SetFocus
        ElseIf tbs.TabEnabled(3) Then
            tbs.Tab = 3
            cmdOpen(3).SetFocus
        ElseIf tbs.TabEnabled(4) Then
            tbs.Tab = 4
            lvw.SetFocus
        Else
            cmdOK.SetFocus
        End If
    End If
End Sub


Private Sub DisableObject()
    VisualTxt.Enabled = True
    cmdOpen(1).Enabled = True
        
    msf.Enabled = True
    cbo(2).Enabled = True
    cboPos(0).Enabled = True
        
    cmdOpen(3).Enabled = True
    cmdClear(0).Enabled = True
    cboPos(1).Enabled = True
    
    lvw.Enabled = True
    tvw.Enabled = True
    cmdLeft.Enabled = True
    cmdRight.Enabled = True
        
    If tbs.TabEnabled(1) = False Then
        VisualTxt.Enabled = False
        cmdOpen(1).Enabled = False
    End If
    If tbs.TabEnabled(2) = False Then
        msf.Enabled = False
        cbo(2).Enabled = False
        cboPos(0).Enabled = False
    End If
    If tbs.TabEnabled(3) = False Then
        cmdOpen(3).Enabled = False
        cmdClear(0).Enabled = False
        cboPos(1).Enabled = False
    End If
    If tbs.TabEnabled(4) = False Then
        lvw.Enabled = False
        tvw.Enabled = False
        cmdLeft.Enabled = False
        cmdRight.Enabled = False
    End If
End Sub
