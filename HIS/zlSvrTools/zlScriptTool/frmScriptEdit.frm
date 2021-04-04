VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmScriptEdit 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "添加  【公共文件】"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7845
   Icon            =   "frmScriptEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   7845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CheckBox chk 
      Caption         =   "强制替换"
      Height          =   315
      Index           =   2
      Left            =   2130
      TabIndex        =   28
      Top             =   7050
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmd 
      Height          =   300
      Index           =   8
      Left            =   7455
      Picture         =   "frmScriptEdit.frx":6852
      Style           =   1  'Graphical
      TabIndex        =   25
      TabStop         =   0   'False
      ToolTipText     =   "清空"
      Top             =   5520
      Width           =   300
   End
   Begin VB.TextBox txtExplanation 
      ForeColor       =   &H00FF0000&
      Height          =   1425
      Left            =   960
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   24
      Top             =   5505
      Width           =   6420
   End
   Begin VB.CommandButton cmd 
      Height          =   300
      Index           =   7
      Left            =   7455
      Picture         =   "frmScriptEdit.frx":D0A4
      Style           =   1  'Graphical
      TabIndex        =   23
      TabStop         =   0   'False
      ToolTipText     =   "全选"
      Top             =   3540
      Width           =   300
   End
   Begin VB.CommandButton cmd 
      Height          =   300
      Index           =   6
      Left            =   7455
      Picture         =   "frmScriptEdit.frx":138F6
      Style           =   1  'Graphical
      TabIndex        =   22
      TabStop         =   0   'False
      ToolTipText     =   "反选"
      Top             =   3930
      Width           =   300
   End
   Begin VB.CommandButton cmd 
      Height          =   300
      Index           =   5
      Left            =   7455
      Picture         =   "frmScriptEdit.frx":1A148
      Style           =   1  'Graphical
      TabIndex        =   21
      TabStop         =   0   'False
      ToolTipText     =   "全清"
      Top             =   4320
      Width           =   300
   End
   Begin VB.CheckBox chk 
      Caption         =   "应用于所有系统"
      Height          =   315
      Index           =   1
      Left            =   3255
      TabIndex        =   18
      Top             =   7050
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.CheckBox chk 
      Caption         =   "升级注册"
      Height          =   315
      Index           =   0
      Left            =   960
      TabIndex        =   17
      Top             =   7050
      Width           =   1095
   End
   Begin VB.TextBox txtVision 
      Height          =   285
      Left            =   4530
      TabIndex        =   16
      Top             =   1215
      Width           =   2835
   End
   Begin VB.CommandButton cmd 
      Height          =   300
      Index           =   4
      Left            =   7425
      Picture         =   "frmScriptEdit.frx":2099A
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "全清"
      Top             =   2550
      Width           =   300
   End
   Begin VB.CommandButton cmd 
      Height          =   300
      Index           =   3
      Left            =   7425
      Picture         =   "frmScriptEdit.frx":271EC
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "反选"
      Top             =   2160
      Width           =   300
   End
   Begin VB.CommandButton cmd 
      Height          =   300
      Index           =   2
      Left            =   7425
      Picture         =   "frmScriptEdit.frx":2DA3E
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "全选"
      Top             =   1770
      Width           =   300
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   6585
      TabIndex        =   10
      Top             =   7035
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "保存(&S)"
      Height          =   350
      Left            =   5370
      TabIndex        =   9
      Top             =   7035
      Width           =   1100
   End
   Begin VB.ComboBox cbo 
      Height          =   300
      Index           =   1
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1215
      Width           =   2625
   End
   Begin VB.CommandButton cmd 
      Height          =   300
      Index           =   1
      Left            =   7425
      Picture         =   "frmScriptEdit.frx":34290
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "选择位置"
      Top             =   690
      Width           =   300
   End
   Begin VB.ComboBox cbo 
      Height          =   300
      Index           =   0
      ItemData        =   "frmScriptEdit.frx":3AAE2
      Left            =   960
      List            =   "frmScriptEdit.frx":3AAE4
      TabIndex        =   4
      Top             =   690
      Width           =   6405
   End
   Begin VB.CommandButton cmd 
      Height          =   300
      Index           =   0
      Left            =   7425
      Picture         =   "frmScriptEdit.frx":3AAE6
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "选择文件"
      Top             =   180
      Width           =   300
   End
   Begin VB.TextBox txtFilePath 
      Height          =   300
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   180
      Width           =   6405
   End
   Begin MSComctlLib.ListView lvwSys 
      Height          =   1635
      Left            =   930
      TabIndex        =   14
      Top             =   1770
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   2884
      View            =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "列名"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "编号"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComDlg.CommonDialog Cdlg 
      Left            =   0
      Top             =   2265
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfCom 
      Height          =   1830
      Left            =   975
      TabIndex        =   20
      Top             =   3540
      Width           =   6405
      _cx             =   11298
      _cy             =   3228
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
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   14737632
      GridColor       =   -2147483638
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   20
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmScriptEdit.frx":41338
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
      Editable        =   1
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
      AccessibleRole  =   30
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   1
      Left            =   975
      TabIndex        =   27
      Top             =   495
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "文件说明"
      Height          =   180
      Index           =   5
      Left            =   165
      TabIndex        =   26
      Top             =   5520
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "业务部件"
      Height          =   180
      Index           =   4
      Left            =   165
      TabIndex        =   19
      Top             =   3540
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "版本号"
      Height          =   180
      Index           =   3
      Left            =   3885
      TabIndex        =   15
      Top             =   1275
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "所属系统"
      Height          =   180
      Index           =   2
      Left            =   135
      TabIndex        =   8
      Top             =   1785
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "部件类型"
      Height          =   180
      Index           =   1
      Left            =   135
      TabIndex        =   7
      Top             =   1270
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "安装路径"
      Height          =   180
      Index           =   0
      Left            =   135
      TabIndex        =   3
      Top             =   755
      Width           =   720
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      Caption         =   "文件位置"
      Height          =   180
      Index           =   0
      Left            =   135
      TabIndex        =   0
      Top             =   240
      Width           =   720
   End
End
Attribute VB_Name = "frmScriptEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private m_blnModed              As Boolean
Private m_str方式               As String
Private m_strNum               As String
Private m_strPathJY             As String
Private m_strEditDate           As String
Private m_lngCurRow             As Long
Private m_strCurFileName        As String
Private mstr序号                As String

Public Property Get Moded() As Boolean
   Moded = m_blnModed
End Property

Public Property Let Moded(ByVal blnModed As Boolean)
    m_blnModed = blnModed
End Property

Private Sub cbo_Click(Index As Integer)
    Dim i As Long
    On Error GoTo errH
    Select Case Index
    Case 0
        
    Case 1
        If m_str方式 = "新增" Then
            If cbo(1).Text = "公共部件" Then
                Call cmd_Click(2)
            Else
                Me.Caption = "添加【" & cbo(1).Text & "】"
                For i = 1 To lvwSys.ListItems.Count
                    If lvwSys.ListItems.Item(i).SubItems(1) = m_strNum Then
                        lvwSys.ListItems.Item(i).Checked = True
                    Else
                        lvwSys.ListItems.Item(i).Checked = False
                    End If
                Next
            End If
        End If
        
        If cbo(1).Text = "系统文件" Then
            chk(2).Visible = True
        Else
            chk(2).Visible = False
        End If
    End Select
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub chk_Click(Index As Integer)
    Dim i As Integer
    Select Case Index
    Case 0 '升级注册
        
        
    Case 1 '应用于所有系统
        If chk(1).Value = 1 Then
            cbo(1).Enabled = False
            lvwSys.Enabled = False
            For i = 1 To lvwSys.ListItems.Count
                lvwSys.ListItems.Item(i).Checked = True
            Next
        Else
            For i = 1 To lvwSys.ListItems.Count
                If lvwSys.ListItems.Item(i).SubItems(1) = m_strNum Then
                    lvwSys.ListItems.Item(i).Checked = True
                Else
                    lvwSys.ListItems.Item(i).Checked = False
                End If
            Next
            cbo(1).Enabled = True
            lvwSys.Enabled = True
            
        End If
    Case 2 '强制覆盖
        
        
    End Select
End Sub

Private Sub cmd_Click(Index As Integer)
    Dim i As Long
    Dim strFilter   As String
    Dim strPath     As String
    Dim strSavePath As String
    Select Case Index
    Case 0 '选择文件
        Dim m_item As MSComctlLib.ListItem
        strFilter = "所有文件" & "|" & "*.*" & "|" & "DLL文件" & "|" & "*.DLL" & "|" & "OCX文件" & "|" & "*.OCX"
        On Error GoTo err
          Cdlg.filename = ""
          Cdlg.DialogTitle = "选择文件"
          Cdlg.MaxFileSize = 8192
          Cdlg.CancelError = True
          Cdlg.InitDir = m_strPathJY
          Cdlg.filename = ""
          Cdlg.Filter = strFilter
          Cdlg.Flags = cdlOFNExplorer
          Cdlg.ShowOpen
          If Cdlg.filename <> "" Then
            strPath = Cdlg.filename
          End If
err:
          If err.Number = cdlCancel Then
             err.Clear
             strPath = ""
          End If
          
          If Len(strPath) Then
            txtFilePath.Text = strPath
            txtVision.Text = GetCommpentVersion(strPath)
            '检查文件是否存在并设置所属系统
            If CheckFile(strPath) Then
                lbl(1).Caption = "该文件已存在!"
            Else
                lbl(1).Caption = ""
            End If
            '分析文件路径类型
            Call AnalyzeFile(strPath)
            
            
            '分析后缀
            If (UCase(Right(strPath, 3))) = "DLL" Or (UCase(Right(strPath, 3))) = "OCX" Then
                chk(0).Value = 1
            Else
                chk(0).Value = 0
            End If
            
          End If
    Case 1 '选择位置
        strSavePath = vbGetBrowseDirectory(Me)
        If strSavePath = "" Then
            Exit Sub
        Else
            cbo(0).Text = strSavePath
        End If
    Case 2 '全选
        For i = 1 To lvwSys.ListItems.Count
            lvwSys.ListItems.Item(i).Checked = True
        Next
    Case 3 '反选
        For i = 1 To lvwSys.ListItems.Count
          If lvwSys.ListItems.Item(i).Checked Then
            lvwSys.ListItems.Item(i).Checked = False
          Else
            lvwSys.ListItems.Item(i).Checked = True
          End If
        Next
        
    Case 4 '全清
        For i = 1 To lvwSys.ListItems.Count
            lvwSys.ListItems.Item(i).Checked = False
        Next
    Case 5 '清空引用文件
        Call StandardAllDel
    Case 6 '删除应用文件
        Call StandardDel
    Case 7 '添加应用文件
        Call AddFile
    Case 8 '清空说明
        txtExplanation.Text = ""
    End Select
End Sub

'==============================================================================
'=功能：取消功能
'==============================================================================
Private Sub cmdCancel_Click()
    If m_str方式 = "新增" Then
        Moded = True
    Else
        Moded = False
    End If
    Unload Me
End Sub

'==============================================================================
'=功能：保存功能
'==============================================================================
Private Sub cmdOK_Click()
    Dim i As Long
    Dim blnSelect  As Boolean
    Dim lngTypeNum As Long
    Dim strPath    As String
    Dim ret        As Long
    On Error GoTo errH
    If txtFilePath = "" Then
        MsgBox "请选择文件!", vbInformation, "提示"
        txtFilePath.SetFocus
        Exit Sub
    End If
    
    If cbo(0) = "" Then
        MsgBox "请选择存放位置!", vbInformation, "提示"
        cbo(0).SetFocus
        Exit Sub
    End If
    
    For i = 1 To lvwSys.ListItems.Count
        If lvwSys.ListItems.Item(i).Checked Then
            blnSelect = True
            Exit For
        End If
    Next
    
    If Len(txtExplanation.Text) > 1900 Then
        MsgBox "文件说明请不要超过2000个字符!", vbInformation, "提示"
        txtExplanation.SetFocus
        Exit Sub
    End If
    
    If blnSelect = False Then
       MsgBox "请选择系统编号!", vbInformation, "提示"
       lvwSys.SetFocus
       Exit Sub
    End If
    
    strPath = cbo(0).Text
    lngTypeNum = cbo(1).ItemData(cbo(1).ListIndex)
 
    
    If SaveDate(txtFilePath, lngTypeNum, strPath) Then
        If m_str方式 = "新增" Then
            ret = MsgBox("是否继续添加?", vbQuestion + vbYesNo, "提示")
            If ret = vbYes Then
                txtFilePath.Text = ""
                txtFilePath.SetFocus
                lbl(1).Caption = "请选择文件!"
                If chk(1).Value = 0 Then
                    For i = 1 To lvwSys.ListItems.Count
                        If lvwSys.ListItems.Item(i).SubItems(1) = m_strNum Then
                            lvwSys.ListItems.Item(i).Checked = True
                        Else
                            lvwSys.ListItems.Item(i).Checked = False
                        End If
                    Next
                End If
            
                Exit Sub
            Else
                Call SaveSetting("zlSvrStudio", "parameter", "Path", cbo(0).Text)
                Call SaveSetting("zlSvrStudio", "parameter", "Type", cbo(1).Text)
                Moded = True
                Unload Me
            End If
        Else
            Call SaveSetting("zlSvrStudio", "parameter", "Path", cbo(0).Text)
            Call SaveSetting("zlSvrStudio", "parameter", "Type", cbo(1).Text)
            Moded = True
            Unload Me
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：页面初始化
'==============================================================================
Private Sub Form_Load()
    On Error GoTo errH

    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能：公共接口函数：用于传入初始化参数:ID '方式为插入，且ID存在，则在ID值前节点插入。
'==============================================================================
Public Sub ShowForm(方式 As String, ByVal 类型名称 As String, ByVal 文件名称 As String, ByVal 所属系统 As String, ByVal 系统号 As String, ByVal 版本号 As String, ByVal 安装路径 As String, ByVal 修改日期 As String, ByVal 所属系统New As String, ByVal 文件说明 As String, ByVal 引用文件 As String, ByVal 自动注册 As Boolean, ByVal 强制覆盖 As Boolean, ByVal 序号 As String)
    On Error GoTo errH
    Dim strPath As String
    Dim strType As String
    m_str方式 = 方式
    m_strNum = 系统号
    
    If 方式 = "新增" Then
        If 序号 <> "0" Then
            mstr序号 = 序号
        Else
            mstr序号 = "0"
        End If
        
        lbl(0).Caption = "文件位置"
        Call FillCboPath
        Call FillCboType
        Call ShowRowName
        Me.Caption = "添加" & "【" & 类型名称 & "】"
        
        '还原上次选择的值
        cmd(0).Enabled = True
        txtFilePath.Enabled = True
        strPath = GetSetting("zlSvrStudio", "parameter", "Path")
        strType = GetSetting("zlSvrStudio", "parameter", "Type")
        If strPath <> "" Then
            cbo(0).Text = strPath
        End If
        
        If strType <> "" Then
            cbo(1).Text = strType
        End If
        txtExplanation.Text = ""
        chk(0).Value = 0
        chk(2).Value = 0
        
        Call initvsfCom
    
    Else
        mstr序号 = "0"
        lbl(0).Caption = "文件名称"
        Call FillCboPath
        Call FillCboType
        Call ShowRowName
        Me.Caption = "修改" & "【" & 类型名称 & "】"
        
        cmd(0).Enabled = False
        txtFilePath.Enabled = False
        txtFilePath.Text = 文件名称
        
        m_strEditDate = 修改日期
        cbo(0).Text = IIf(安装路径 = "0", "", 安装路径)
        cbo(1).Text = 类型名称
        txtVision.Text = IIf(版本号 = "0", "", 版本号)
        
        Dim i As Integer
        Dim j As Integer
        Dim strArr As Variant
        
        If 所属系统New = "" Then
            For i = 1 To lvwSys.ListItems.Count
                lvwSys.ListItems.Item(i).Checked = True
            Next
        Else
            For i = 1 To lvwSys.ListItems.Count
                lvwSys.ListItems.Item(i).Checked = False
            Next
            
            strArr = Split(所属系统New, ",")
            For i = 0 To UBound(strArr) - 1
                If strArr(i) <> "" Then
                    For j = 1 To lvwSys.ListItems.Count
                        If strArr(i) = lvwSys.ListItems.Item(j).SubItems(1) Then
                            lvwSys.ListItems.Item(j).Checked = True
                            Exit For
                        End If
                    Next
                End If
            Next
        End If
        
'''        Call SetChk(系统参数)
        If 文件说明 = "0" Then
            txtExplanation.Text = ""
        Else
            txtExplanation.Text = 文件说明
        End If
        
        If 自动注册 Then
            chk(0).Value = 1
        Else
            chk(0).Value = 0
        End If
        
        If 强制覆盖 Then
            chk(2).Value = 1
        Else
            chk(2).Value = 0
        End If
        
        Call initvsfCom
        If Len(引用文件) > 0 Then
            Call refvsfCom(引用文件)
        End If
    End If
    
    Me.Show 1
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'填充安装路径默认值
Private Sub FillCboPath()
    On Error GoTo errH
    With cbo(0)
        .Clear
        
'        .AddItem "[Apply]"
'        .ItemData(.NewIndex) = 0
        .AddItem "[AppSoft]"
        .ItemData(.NewIndex) = 0
        .AddItem "[System]"
        .ItemData(.NewIndex) = 1
        .AddItem "[Help]"
        .ItemData(.NewIndex) = 2
        .AddItem "[Public]"
        .ItemData(.NewIndex) = 3
'        .AddItem "[附加文件]"
'        .ItemData(.NewIndex) = 4
'        .AddItem "[PacsList]"
'        .ItemData(.NewIndex) = 5
'        .AddItem "[InSureNew]"
'        .ItemData(.NewIndex) = 6
    
        .ListIndex = 0
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'填充文件类型默认值
Private Sub FillCboType()
    On Error GoTo errH
    With cbo(1)
        .Clear
        
        .AddItem "公共部件"
        .ItemData(.NewIndex) = 0
        .AddItem "应用部件"
        .ItemData(.NewIndex) = 1
        .AddItem "帮助文件"
        .ItemData(.NewIndex) = 2
        .AddItem "其它文件"
        .ItemData(.NewIndex) = 3
        .AddItem "三方部件"
        .ItemData(.NewIndex) = 4
        .AddItem "系统文件"
        .ItemData(.NewIndex) = 5
        
        .ListIndex = 0
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


''显示指定表的所有列名
Private Sub ShowRowName()
  Dim strSQL As String, rs As ADODB.Recordset
  Dim m_list As MSComctlLib.ListItem
  Dim i As Integer
  Dim str编号 As String
  On Error GoTo errH
  lvwSys.ListItems.Clear
  strSQL = "select 名称,编号 from zlSystems order by 编号"
  Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, 30, 100)
  If rs.RecordCount > 0 Then
    rs.MoveFirst
    Do Until rs.EOF
      str编号 = NVL(rs!编号) \ 100
      Set m_list = lvwSys.ListItems.Add(, , "[" & str编号 & "]" & NVL(rs!名称))
          m_list.SubItems(1) = str编号
      rs.MoveNext
    Loop
  End If
  Exit Sub
errH:
  If ErrCenter() = 1 Then Resume
  Call SaveErrLog
End Sub

'保存数据
Private Function SaveDate(ByVal strFilePath As String, ByVal lngTypeNum As Long, ByVal strPath As String) As Boolean
    Dim rs          As ADODB.Recordset
    Dim rsMaxID     As ADODB.Recordset
    Dim rsShell     As ADODB.Recordset
    Dim strSQL      As String
    Dim strName     As String '名称
    Dim strVision   As String '版本号
    Dim strEditDate As String '修改日期
    Dim ret         As Long
    Dim strArr      As Variant
    Dim lng版本号   As Double
    Dim i           As Long
    Dim strMax序号  As String '最大序号
    Dim strCurSelectSys As String
    Dim dateEdit    As Date  '修改日期
    Dim lngSelectNum As Long '选择数
    Dim bln所属系统 As Boolean
    
    
    Dim str所属系统 As String '所属系统
'    Dim bln注册文件 As Boolean '注册文件
'    Dim bln应用所有 As Boolean '应用所有
'    Dim str注册文件 As String
'    Dim str应用所有 As String
'    Dim str组合参数 As String '组合成参数串保存在数据库中
    Dim str文件说明 As String '文件说明
    Dim str引用文件 As String '引用文件
    Dim byt自动注册 As Byte
    Dim byt强制覆盖 As Byte
    Dim dateJoin As Date '加入日期
    
    On Error GoTo errH
    lngSelectNum = 0
    strName = UCase(GetFileName(strFilePath, , True))
    strSQL = "select 文件名,所属系统 from zlFilesUpgrade where upper(文件名) = upper('" & strName & "') "
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)

    
    '获得最大序号
    If m_str方式 = "新增" Then
        If mstr序号 <> "0" Then
            strMax序号 = CLng(Val(mstr序号))
        Else
            strSQL = "select max(to_number(序号)) as 序号 from  zlFilesUpgrade"
            Set rsMaxID = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
            If rsMaxID.RecordCount > 0 Then
                strMax序号 = CLng(NVL(rsMaxID!序号, 0))
            Else
                strMax序号 = 1
            End If
        End If
        
        '获得修改日期
        dateEdit = Format(FileDateTime(strFilePath), "yyyy-MM-dd hh:mm:ss")
    Else
        dateEdit = Format(m_strEditDate, "yyyy-mm-dd hh:mm:ss")
    End If
    

    '更新文件
    '组合存储版本号
    strVision = txtVision.Text
    If strVision <> "" Then
        strArr = Split(strVision, ".")
        lng版本号 = strArr(0) * 10 ^ 8 + strArr(1) * 10 ^ 4 + strArr(2)
        strVision = lng版本号
    End If
    
    
    '当前选择的所属系统
    With lvwSys
        For i = 1 To .ListItems.Count
            If .ListItems.Item(i).Checked Then
                lngSelectNum = lngSelectNum + 1
                If strCurSelectSys = "" Then
                    strCurSelectSys = "," & .ListItems.Item(i).SubItems(1)
                Else
                    strCurSelectSys = strCurSelectSys & "," & .ListItems.Item(i).SubItems(1)
                End If
            End If
        Next
        If Len(strCurSelectSys) <> 0 Then
            strCurSelectSys = strCurSelectSys & ","
        End If
        If lngSelectNum = .ListItems.Count Then
            bln所属系统 = True
        Else
            bln所属系统 = False
        End If
    End With
    
    
'    bln注册文件 = chk(0).Value
'    bln应用所有 = chk(1).Value
    
    str文件说明 = txtExplanation.Text
    str引用文件 = getFiles
    byt自动注册 = IIf(chk(0).Value = 0, 0, 1)
    byt强制覆盖 = IIf(chk(2).Value = 0, 0, 1)
    
    dateJoin = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:MM:SS")
    If rs.RecordCount > 0 Then
            
            If bln所属系统 Then
                str所属系统 = ""
            Else
                If NVL(rs!所属系统) <> "" Then
                    str所属系统 = rs!所属系统
                    str所属系统 = GetUnionSysNum(str所属系统, strCurSelectSys)
                Else
                    str所属系统 = strCurSelectSys
                End If
            End If
            
'            str注册文件 = ";Z" & IIf(bln注册文件 = True, 1, 0)
'            str应用所有 = ";Y" & IIf(bln应用所有 = True, 1, 0)
'            str组合参数 = str注册文件 & str应用所有
        
        
            '更新SQL执行
'            gcnOracle.BeginTrans
            strSQL = "update Zlfilesupgrade set 文件类型='" & lngTypeNum & "',版本号='" & strVision & "',业务部件='" & str引用文件 & "',所属系统='" & str所属系统 & "',安装路径='" & strPath & "'" & _
            ",修改日期=to_date('" & CStr(dateEdit) & "','yyyy-mm-dd hh24:mi:ss'),文件说明='" & str文件说明 & "',强制覆盖=" & byt强制覆盖 & ",自动注册=" & byt自动注册 & " where upper(文件名)='" & strName & "'"
            gcnOracle.Execute strSQL
'            gcnOracle.CommitTrans
            
            
            '            Set rsShell = gmobjCommon.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngTypeNum, strVision, str所属系统, strPath, CStr(dateEdit), strName)
            SaveDate = True
            Exit Function
    Else
        '插入SQL执行
        If bln所属系统 Then
            str所属系统 = ""
        Else
            str所属系统 = strCurSelectSys
        End If
'        str注册文件 = ";Z" & IIf(bln注册文件 = True, 1, 0)
'        str应用所有 = ";Y" & IIf(bln应用所有 = True, 1, 0)
'        str组合参数 = str注册文件 & str应用所有
        If mstr序号 <> "0" Then
            strSQL = "update zlfilesupgrade set 序号= 序号+1 Where 序号>" & Val(strMax序号)
            gcnOracle.Execute strSQL
        End If
   
'        gcnOracle.BeginTrans
        strSQL = "insert into zlFilesUpgrade (序号,文件类型,文件名,版本号,修改日期,业务部件,所属系统,安装路径,文件说明,强制覆盖,自动注册,加入日期) values (" & strMax序号 + 1 & "," & lngTypeNum & "," & _
        "'" & strName & "','" & strVision & "',to_date('" & CStr(dateEdit) & "','yyyy-mm-dd hh24:mi:ss'),'" & str引用文件 & "','" & str所属系统 & "','" & strPath & "','" & str文件说明 & "'," & byt强制覆盖 & " ," & byt自动注册 & ",to_date('" & CStr(dateJoin) & "','yyyy-mm-dd hh24:mi:ss'))"
        
        gcnOracle.Execute strSQL
        
   
        
'        gcnOracle.CommitTrans
'        Set rsShell = gmobjCommon.zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strMax序号 + 1, lngTypeNum, strName, strVision, CStr(dateEdit), str所属系统, strPath)
'        gcnOracle.CommitTrans
        
        
        SaveDate = True
        Exit Function
    End If
    
    Exit Function
errH:
'    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function GetUnionSysNum(ByVal str所属系统 As String, ByVal strCurSelectSys As String) As String
    On Error GoTo errH
    Dim strArr As Variant
    Dim i As Integer
    
    Dim strTemp As String
    strTemp = ""
    strArr = Split(strCurSelectSys, ",")
    For i = 0 To UBound(strArr) - 1
        If strArr(i) <> "" Then
            If InStrRev(strCurSelectSys, "," & strArr(i) & ",", 1) = 0 Then
                If strCurSelectSys <> "," & strArr(i) & "," Then
                    strTemp = strTemp & "," & strArr(i)
                End If
            End If
        End If
    Next
    
    If strTemp <> "" Then
        strTemp = strTemp & ","
        GetUnionSysNum = strTemp
        
    Else
        GetUnionSysNum = strCurSelectSys
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'==============================================================================
'=功能： 检查文件是否存在于数据库中，存在并列出它的所属!
'==============================================================================
Private Function CheckFile(ByVal strFile As String) As Boolean
    On Error GoTo errH
    Dim rs          As ADODB.Recordset
    Dim strSQL      As String
    Dim str所属     As String
    Dim strArr      As Variant
    Dim i           As Integer
    Dim j           As Integer
    
    strFile = UCase(GetFileName(strFile, , True))
    strSQL = "select 所属系统,业务部件,安装路径,文件说明,1 from zlFilesUpgrade where upper(文件名) = upper([1])"
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strFile)
    
    If rs.RecordCount = 1 Then
        rs.MoveFirst
        
        For i = 1 To lvwSys.ListItems.Count
            lvwSys.ListItems.Item(i).Checked = False
        Next
        
        str所属 = NVL(rs!所属系统, "")
        strArr = Split(str所属, ",")
        For i = 0 To UBound(strArr) - 1
            If strArr(i) <> "" Then
                For j = 1 To lvwSys.ListItems.Count
                    If strArr(i) = lvwSys.ListItems.Item(j).SubItems(1) Then
                        lvwSys.ListItems.Item(j).Checked = True
                        Exit For
                    End If
                Next
            End If
        Next
        
        
        
        cbo(0).Text = NVL(rs!安装路径, "")
'        Call SetChk(NVL(rs!系统参数, ""))
     
        txtExplanation.Text = NVL(rs!文件说明, "")
     
        If Len(NVL(rs!业务部件, "")) > 0 Then
            Call refvsfCom(NVL(rs!业务部件, ""))
        End If
        
        CheckFile = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'==============================================================================
'=功能： 分析文件的类型
'==============================================================================
Private Sub AnalyzeFile(ByVal strFile As String)
    On Error GoTo errH
    Dim i As Integer
    Dim strWinSystemPath As String
    Dim strWinPath       As String
    Dim strHelp          As String
    Dim strApp           As String
    
    strFile = UCase(strFile)
    strWinSystemPath = UCase(GetWinSystemPath())
    strWinPath = UCase(GetWinPath())
'    strMainPan = UCase(Left(strWinPath, 1))
    strHelp = UCase(strWinPath & "\HELP")


    If InStrRev(strFile, strWinSystemPath, -1, vbTextCompare) > 0 Then
        cbo(0).ListIndex = 1
    ElseIf InStrRev(strFile, strHelp, -1, vbTextCompare) > 0 Then
        cbo(0).ListIndex = 2
    ElseIf InStrRev(strFile, "\APPSOFT\", -1, vbTextCompare) > 0 Then
        strApp = GetAppSoft(strFile)
        If strApp = "" Then
            cbo(0).ListIndex = 0
        Else
            cbo(0).Text = "[APPSOFT]\" & strApp
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'''Private Sub SetChk(ByVal strSysOption As String)
'''    On Error GoTo errH
'''    If strSysOption = "" Or strSysOption = "0" Then Exit Sub
'''    Dim i As Integer, j As Integer
'''    i = InStrRev(strSysOption, "Z", -1)
'''    If i > 0 Then
'''        chk(0).Value = Right(Left(strSysOption, i + 1), 1)
'''    End If
'''
'''    j = InStrRev(strSysOption, "Y", -1)
'''    If j > 0 Then
'''        chk(1).Value = Right(Left(strSysOption, j + 1), 1)
'''    End If
'''    Exit Sub
'''errH:
'''    If ErrCenter() = 1 Then Resume
'''    Call SaveErrLog
'''End Sub

'==============================================================================
'=功能： 初始化VSFCom
'==============================================================================
Private Sub initvsfCom()
    On Error GoTo errH
    With vsfCom
        .Tag = ""
'        .Redraw = flexRDNone
        .Rows = 6
        .Clear
        .Cols = 2
        .Cell(flexcpText, 0, 0) = "序号"
        .Cell(flexcpAlignment, 0, 0) = flexAlignCenterCenter
        .ColWidth(0) = 500
        .Cell(flexcpText, 0, 1) = "文件名"
        .Cell(flexcpAlignment, 0, 1) = flexAlignCenterCenter
        .ColWidth(1) = 5000
'        .Cell(flexcpText, 0, 2) = "版本号"
'        .Cell(flexcpAlignment, 0, 2) = flexAlignCenterCenter
'        .ColWidth(2) = 1000
'        .Cell(flexcpText, 0, 3) = "修改日期"
'        .Cell(flexcpAlignment, 0, 2) = flexAlignCenterCenter
'        .ColWidth(3) = 1000
        '自动换行
        .WordWrap = True
        '合并单元格
        .MergeCells = 0
        .MergeCol(.ColIndex("文件类型")) = True
        .MergeCol(.ColIndex("文件名")) = True
        '隐藏单元格
        '行高设置
        .RowHeightMin = 300
        '最大宽度设置
        .ColWidthMax = 4000
        '自动适应行高、列宽
        .AutoSizeMode = flexAutoSizeRowHeight
'        .AutoSize .ColIndex("文件名")
'        .SelectionMode = flexSelectionListBox
        .AllowBigSelection = False
        .Redraw = flexRDBuffered
        
    End With
    Exit Sub
errH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog

End Sub

'刷新vsfCom
Private Sub refvsfCom(ByVal strFiles As String)
    On Error GoTo errH
    Dim i As Long
    Dim iNum As String
    Dim strTemp() As String
    Call initvsfCom
    If strFiles = "" Then Exit Sub
    strTemp = Split(strFiles, ",")
    
    With vsfCom
        .Rows = UBound(strTemp) + 2
        For i = 0 To UBound(strTemp)
            .Cell(flexcpText, i + 1, 0) = i + 1
            .Cell(flexcpAlignment, i + 1, 0) = flexAlignLeftCenter
            .Cell(flexcpText, i + 1, 1) = strTemp(i)
            .Cell(flexcpAlignment, i + 1, 1) = flexAlignLeftCenter
        Next
        
'        '自动换行
'        .WordWrap = True
'        '合并单元格
'        .MergeCells = 2
'        .MergeCol(.ColIndex("文件类型")) = True
'        .MergeCol(.ColIndex("文件名")) = True
'        '隐藏单元格
'        .ColWidth(.ColIndex("类型ID")) = 0
'        '行高设置
'        .RowHeightMin = 300
'        '最大宽度设置
'        .ColWidthMax = 7000
'        '自动适应行高、列宽
'        .AutoSizeMode = flexAutoSizeRowHeight
'        .AutoSize .ColIndex("业务部件")
'        .SelectionMode = flexSelectionListBox
'        .AllowBigSelection = False
'        .Redraw = flexRDBuffered
    End With
    
    Exit Sub
errH:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
    
End Sub

Private Sub AddFile()
    Dim strFiles As String
    On Error GoTo errH
    
        strFiles = getFiles
        With frmEditFile
            .intItemFile = strFiles
            .intStrFile = txtFilePath.Text
            .strType = "0,1,2,3,4"
            .Show vbModal
            
            Call refvsfCom(.intItemFile)
         
        End With
        Set frmEditFile = Nothing
        Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtExplanation_GotFocus()
    txtExplanation.BackColor = &HC0FFC0
End Sub

Private Sub txtExplanation_LostFocus()
    txtExplanation.BackColor = &H80000005
End Sub

Private Function getFiles() As String
    On Error GoTo errH
    Dim strTemp As String
    Dim i As Long
    strTemp = ""
    For i = 1 To vsfCom.Rows - 1
        If strTemp = "" Then
            If vsfCom.Cell(flexcpText, i, 1) <> "" Then
                strTemp = vsfCom.Cell(flexcpText, i, 1) & ","
            End If
        Else
            If vsfCom.Cell(flexcpText, i, 1) <> "" Then
                strTemp = strTemp & vsfCom.Cell(flexcpText, i, 1) & ","
            End If
        End If
    Next
    
    If Len(strTemp) <> 0 Then
        If Right(strTemp, 1) = "," Then
            strTemp = Left(strTemp, Len(strTemp) - 1)
        End If
        getFiles = strTemp
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

'==============================================================================
'=删除引用
'==============================================================================
Private Sub StandardDel()
    On Error GoTo errH
    Dim lngRow As Long
    Dim strSelectFile As String
    Dim strFiles As String

    If m_strCurFileName = "" Then Exit Sub
    strFiles = getFiles
    If strFiles <> "" Then
        lngRow = vsfCom.FindRow(CStr(m_strCurFileName), , 1)
        

        strFiles = strFiles & ","
        strFiles = Replace(strFiles, m_strCurFileName & ",", "")
        If Right(strFiles, 1) = "," Then
            strFiles = Left(strFiles, Len(strFiles) - 1)
        End If
        Call refvsfCom(strFiles)
        
        If lngRow <> -1 Then
            If lngRow >= 2 Then
              vsfCom.Select lngRow - 1, 1
              vsfCom.ShowCell lngRow - 1, 1
            End If
        End If
    End If
   
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=删除所有引用
'==============================================================================
Private Sub StandardAllDel()
    On Error GoTo errH
    Call initvsfCom
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


'==============================================================================
'=功能： 网格行列变化时更新基本信息
'==============================================================================
Private Sub vsfcom_RowColChange()
    On Error GoTo errH
    Call vsfcom_SelChange
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'==============================================================================
'=功能： 网格选择行列变化时更新基本信息
'==============================================================================
Private Sub vsfcom_SelChange()
    Dim lngID       As Long
    On Error GoTo errH
    
'    fgMain.WallPaper = imgBG_fg(1).Picture
    m_lngCurRow = vsfCom.Row
    If m_lngCurRow = 0 Then Exit Sub
    m_strCurFileName = IIf(Len(vsfCom.Cell(flexcpText, m_lngCurRow, 1)) = 0, "", vsfCom.Cell(flexcpText, m_lngCurRow, 1))   '获取ID
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function GetAppSoft(ByVal strFile As String) As String
    On Error Resume Next
    Dim i As Integer
    Dim j As Integer
    Dim strTemp As String
    i = InStrRev(strFile, "\APPSOFT\", -1)
    strTemp = Right(strFile, Len(strFile) - (i + 8))
    i = InStrRev(strTemp, "\", -1)
    If i > 0 Then
        GetAppSoft = Left(strTemp, i)
    Else
        GetAppSoft = ""
    End If
End Function
